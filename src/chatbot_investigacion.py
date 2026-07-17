"""
Chatbot analítico de investigación (reemplazo de "Panorama General").

Módulo autocontenido: toda la lógica (consultas a la BD, herramientas del
LLM, worker de Ollama, UI de chat + gráficas) vive acá. La integración con
`main_10.py` es mínima (una pestaña nueva + un callback de navegación) --
ver VentanaPrincipal._cambiar_pestana_por_nombre.

Requiere Ollama corriendo localmente (http://localhost:11434) con un modelo
que soporte "tool calling" (default: qwen2.5:7b -- `ollama pull qwen2.5:7b`).

Los datos de GrupLAC (cumplimiento, faltantes) siempre se leen desde
data/cache/verificacion_faltantes.json vía cargar_df_faltantes(), el mismo
caché que usa el panel Cumplimiento de Seguimiento Grupos -- ese caché ya se
genera contra la carpeta 'data/reporte excel_<fecha>' más reciente, así que
nunca se compara contra un scrape viejo.
"""

from __future__ import annotations

import json
from collections import defaultdict
from datetime import datetime
from typing import Any, Dict, List, Optional

import numpy as np
import requests
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QComboBox, QHBoxLayout, QLabel, QLineEdit, QPushButton, QSpinBox,
    QSplitter, QTextEdit, QVBoxLayout, QWidget,
)

from views.vista_seguimiento_grupos import cargar_df_faltantes

MODELO_DEFAULT = "qwen2.5:7b"


# =============================================================================
# CONSULTOR DE BASE DE DATOS + CACHÉ GRUPLAC
# =============================================================================

class ConsultorAnalitico:
    """Consultas de solo lectura sobre la BD interna y el caché de
    verificación GrupLAC, en forma de diccionarios simples (listos para
    volcar a JSON y pasarle al LLM)."""

    def __init__(self, db_manager):
        self.db = db_manager
        self.cursor = self.db.conn.cursor()

    # ── Grupos / integrantes ────────────────────────────────────────────
    def listar_grupos(self, filtro_texto: Optional[str] = None) -> List[Dict]:
        query = '''
            SELECT g.grupo, COUNT(DISTINCT g.cedula) as num_integrantes, g.facultad
            FROM grupos g
            WHERE g.grupo IS NOT NULL AND g.grupo != ''
        '''
        params: list = []
        if filtro_texto:
            query += " AND LOWER(g.grupo) LIKE LOWER(?)"
            params.append(f"%{filtro_texto}%")
        query += " GROUP BY g.grupo ORDER BY g.grupo"

        resultados = self.cursor.execute(query, params).fetchall()
        return [
            {
                'nombre': grupo,
                'integrantes': num_int,
                'facultad': facultad or 'No especificada',
                'estadisticas': self._obtener_estadisticas_grupo(grupo),
            }
            for grupo, num_int, facultad in resultados
        ]

    def _obtener_estadisticas_grupo(self, grupo: str) -> Dict[str, int]:
        stats = {}
        stats['publicaciones'] = self.cursor.execute(
            "SELECT COUNT(*) FROM publicaciones WHERE cedula IN (SELECT cedula FROM grupos WHERE grupo = ?)",
            (grupo,)).fetchone()[0]
        stats['extensiones'] = self.cursor.execute(
            "SELECT COUNT(*) FROM extensiones WHERE cedula IN (SELECT cedula FROM grupos WHERE grupo = ?)",
            (grupo,)).fetchone()[0]
        stats['trabajos_grado'] = self.cursor.execute(
            "SELECT COUNT(*) FROM trabajos_grado WHERE cedula_director IN (SELECT cedula FROM grupos WHERE grupo = ?)",
            (grupo,)).fetchone()[0]
        stats['proyectos'] = self.cursor.execute(
            "SELECT COUNT(*) FROM proyectos WHERE cedula IN (SELECT cedula FROM grupos WHERE grupo = ?)",
            (grupo,)).fetchone()[0]
        return stats

    def obtener_detalle_grupo(self, nombre_grupo: str) -> Dict[str, Any]:
        integrantes = self.cursor.execute('''
            SELECT p.nombre, p.cedula, g.tipo_miembro, p.email
            FROM grupos g JOIN personas p ON g.cedula = p.cedula
            WHERE g.grupo = ? ORDER BY p.nombre
        ''', (nombre_grupo,)).fetchall()

        publicaciones = self.cursor.execute('''
            SELECT titulo, año, tipo, categoria, revista_libro FROM publicaciones
            WHERE cedula IN (SELECT cedula FROM grupos WHERE grupo = ?) AND año >= 2022
            ORDER BY año DESC LIMIT 15
        ''', (nombre_grupo,)).fetchall()

        proyectos = self.cursor.execute('''
            SELECT titulo, año, tipo, estado, codigo_cie FROM proyectos
            WHERE cedula IN (SELECT cedula FROM grupos WHERE grupo = ?)
            ORDER BY año DESC LIMIT 10
        ''', (nombre_grupo,)).fetchall()

        if not integrantes:
            return {"error": f"No se encontró el grupo '{nombre_grupo}'. Verifique el nombre exacto."}

        return {
            'nombre': nombre_grupo,
            'integrantes': [
                {'nombre': i[0], 'cedula': i[1], 'tipo': i[2], 'email': i[3]} for i in integrantes
            ],
            'publicaciones_recientes': [
                {'titulo': p[0], 'año': p[1], 'tipo': p[2], 'categoria': p[3], 'revista': p[4]}
                for p in publicaciones
            ],
            'proyectos': [
                {'titulo': pr[0], 'año': pr[1], 'tipo': pr[2], 'estado': pr[3], 'codigo': pr[4]}
                for pr in proyectos
            ],
            'estadisticas': self._obtener_estadisticas_grupo(nombre_grupo),
        }

    # ── Producción / tendencias ──────────────────────────────────────────
    def analizar_produccion_por_anio(self, anio_inicio: int, anio_fin: int) -> Dict[str, Any]:
        analisis = {'rango': f"{anio_inicio}-{anio_fin}", 'por_tipo': {},
                    'por_anio': defaultdict(lambda: defaultdict(int)), 'total': 0}

        pubs = self.cursor.execute(
            "SELECT año, tipo, COUNT(*) FROM publicaciones WHERE año BETWEEN ? AND ? GROUP BY año, tipo",
            (anio_inicio, anio_fin)).fetchall()
        for anio, tipo, cantidad in pubs:
            analisis['por_anio'][anio]['publicaciones'] += cantidad
            clave = f"Publicaciones {tipo}" if tipo else "Publicaciones"
            analisis['por_tipo'][clave] = analisis['por_tipo'].get(clave, 0) + cantidad
            analisis['total'] += cantidad

        proys = self.cursor.execute(
            "SELECT año, tipo, COUNT(*) FROM proyectos WHERE año BETWEEN ? AND ? GROUP BY año, tipo",
            (anio_inicio, anio_fin)).fetchall()
        for anio, tipo, cantidad in proys:
            analisis['por_anio'][anio]['proyectos'] += cantidad
            clave = f"Proyectos {tipo}" if tipo else "Proyectos"
            analisis['por_tipo'][clave] = analisis['por_tipo'].get(clave, 0) + cantidad
            analisis['total'] += cantidad

        ext = self.cursor.execute(
            "SELECT año, COUNT(*) FROM extensiones WHERE año BETWEEN ? AND ? GROUP BY año",
            (anio_inicio, anio_fin)).fetchall()
        for anio, cantidad in ext:
            analisis['por_anio'][anio]['extensiones'] = cantidad
            analisis['por_tipo']['Extensiones'] = analisis['por_tipo'].get('Extensiones', 0) + cantidad
            analisis['total'] += cantidad

        analisis['por_anio'] = dict(analisis['por_anio'])
        return analisis

    def obtener_top_investigadores(self, limite: int = 10, tipo_producto: str = 'todos') -> List[Dict]:
        if tipo_producto == 'publicaciones':
            query = '''
                SELECT p.nombre, p.cedula, COUNT(*) as total FROM personas p
                JOIN publicaciones pub ON p.cedula = pub.cedula
                GROUP BY p.cedula, p.nombre ORDER BY total DESC LIMIT ?
            '''
        elif tipo_producto == 'proyectos':
            query = '''
                SELECT p.nombre, p.cedula, COUNT(*) as total FROM personas p
                JOIN proyectos pr ON p.cedula = pr.cedula
                GROUP BY p.cedula, p.nombre ORDER BY total DESC LIMIT ?
            '''
        else:
            query = '''
                SELECT p.nombre, p.cedula,
                    (SELECT COUNT(*) FROM publicaciones WHERE cedula = p.cedula) +
                    (SELECT COUNT(*) FROM proyectos WHERE cedula = p.cedula) +
                    (SELECT COUNT(*) FROM extensiones WHERE cedula = p.cedula) +
                    (SELECT COUNT(*) FROM trabajos_grado WHERE cedula_director = p.cedula) as total
                FROM personas p WHERE total > 0 ORDER BY total DESC LIMIT ?
            '''
        resultados = self.cursor.execute(query, (limite,)).fetchall()
        return [
            {'nombre': r[0], 'cedula': r[1], 'total_productos': r[2],
             'grupos': self._obtener_grupos_de_cedula(r[1])}
            for r in resultados
        ]

    def _obtener_grupos_de_cedula(self, cedula: str) -> List[str]:
        grupos = self.cursor.execute(
            "SELECT DISTINCT grupo FROM grupos WHERE cedula = ?", (cedula,)).fetchall()
        return [g[0] for g in grupos if g[0]]

    def comparar_grupos(self, grupo1: str, grupo2: str) -> Dict[str, Any]:
        d1 = self.obtener_detalle_grupo(grupo1)
        d2 = self.obtener_detalle_grupo(grupo2)
        if "error" in d1 or "error" in d2:
            return {"error": d1.get("error") or d2.get("error")}
        return {
            'grupo1': {'nombre': grupo1, 'integrantes': len(d1['integrantes']), 'estadisticas': d1['estadisticas']},
            'grupo2': {'nombre': grupo2, 'integrantes': len(d2['integrantes']), 'estadisticas': d2['estadisticas']},
            'diferencias': {
                'integrantes': len(d1['integrantes']) - len(d2['integrantes']),
                'publicaciones': d1['estadisticas']['publicaciones'] - d2['estadisticas']['publicaciones'],
                'proyectos': d1['estadisticas']['proyectos'] - d2['estadisticas']['proyectos'],
            },
        }

    def buscar_productos_por_palabras_clave(self, palabras: List[str],
                                             tipo: Optional[str] = None,
                                             anio_desde: Optional[int] = None) -> List[Dict]:
        resultados: List[Dict] = []
        palabras = [p for p in palabras if p]
        if not palabras:
            return resultados

        if tipo in (None, 'publicacion'):
            condiciones = " AND ".join("LOWER(titulo) LIKE LOWER(?)" for _ in palabras)
            query = f"SELECT titulo, año, tipo, categoria FROM publicaciones WHERE {condiciones}"
            params = [f"%{p}%" for p in palabras]
            if anio_desde:
                query += " AND año >= ?"
                params.append(anio_desde)
            query += " ORDER BY año DESC LIMIT 20"
            for p in self.cursor.execute(query, params).fetchall():
                resultados.append({'tipo': 'Publicación', 'titulo': p[0], 'año': p[1],
                                    'subtipo': p[2], 'categoria': p[3]})

        if tipo in (None, 'proyecto'):
            condiciones = " AND ".join("LOWER(titulo) LIKE LOWER(?)" for _ in palabras)
            query = f"SELECT titulo, año, tipo, estado FROM proyectos WHERE {condiciones}"
            params = [f"%{p}%" for p in palabras]
            if anio_desde:
                query += " AND año >= ?"
                params.append(anio_desde)
            query += " ORDER BY año DESC LIMIT 20"
            for pr in self.cursor.execute(query, params).fetchall():
                resultados.append({'tipo': 'Proyecto', 'titulo': pr[0], 'año': pr[1],
                                    'subtipo': pr[2], 'estado': pr[3]})
        return resultados

    def obtener_tendencias_investigacion(self, ventana_anios: int = 5) -> Dict[str, Any]:
        anio_actual = datetime.now().year
        anio_inicio = anio_actual - ventana_anios

        categorias = self.cursor.execute('''
            SELECT categoria, COUNT(*) FROM publicaciones
            WHERE año >= ? AND categoria IS NOT NULL AND categoria != ''
            GROUP BY categoria ORDER BY COUNT(*) DESC LIMIT 10
        ''', (anio_inicio,)).fetchall()

        tipos_proyecto = self.cursor.execute('''
            SELECT tipo, COUNT(*) FROM proyectos
            WHERE año >= ? AND tipo IS NOT NULL AND tipo != ''
            GROUP BY tipo ORDER BY COUNT(*) DESC LIMIT 10
        ''', (anio_inicio,)).fetchall()

        evolucion = self.cursor.execute('''
            SELECT año, COUNT(*) FROM publicaciones WHERE año >= ? GROUP BY año ORDER BY año
        ''', (anio_inicio,)).fetchall()

        return {
            'periodo': f"{anio_inicio}-{anio_actual}",
            'categorias_populares': [{'categoria': c[0], 'cantidad': c[1]} for c in categorias],
            'tipos_proyecto_comunes': [{'tipo': t[0], 'cantidad': t[1]} for t in tipos_proyecto],
            'evolucion_anual': [{'año': e[0], 'publicaciones': e[1]} for e in evolucion],
        }

    def buscar_investigador(self, nombre_parcial: str) -> List[Dict]:
        investigadores = self.cursor.execute('''
            SELECT p.nombre, p.cedula, p.email, p.facultad FROM personas p
            WHERE LOWER(p.nombre) LIKE LOWER(?) LIMIT 10
        ''', (f"%{nombre_parcial}%",)).fetchall()
        return [
            {'nombre': inv[0], 'cedula': inv[1], 'email': inv[2], 'facultad': inv[3],
             'grupos': self._obtener_grupos_de_cedula(inv[1]),
             'produccion': self._obtener_estadisticas_investigador(inv[1])}
            for inv in investigadores
        ]

    def _obtener_estadisticas_investigador(self, cedula: str) -> Dict[str, int]:
        stats = {}
        stats['publicaciones'] = self.cursor.execute(
            'SELECT COUNT(*) FROM publicaciones WHERE cedula = ?', (cedula,)).fetchone()[0]
        stats['proyectos'] = self.cursor.execute(
            'SELECT COUNT(*) FROM proyectos WHERE cedula = ?', (cedula,)).fetchone()[0]
        stats['extensiones'] = self.cursor.execute(
            'SELECT COUNT(*) FROM extensiones WHERE cedula = ?', (cedula,)).fetchone()[0]
        stats['trabajos_dirigidos'] = self.cursor.execute(
            'SELECT COUNT(*) FROM trabajos_grado WHERE cedula_director = ?', (cedula,)).fetchone()[0]
        return stats

    # ── GrupLAC (caché de verificación) ─────────────────────────────────
    def cumplimiento_grupo(self, grupo: str) -> Dict[str, Any]:
        """% de cumplimiento (confirmados / (confirmados+faltantes)), misma
        fórmula que el panel Cumplimiento de Seguimiento Grupos."""
        df, mensaje, ts = cargar_df_faltantes()
        if df is None:
            return {"error": mensaje}
        df_g = df[df["grupo_original"] == grupo]
        if df_g.empty:
            return {"error": f"No hay datos de verificación GrupLAC para '{grupo}'. "
                              "Verifique el nombre exacto o corra 'Verificar contra GrupLAC' "
                              "en Seguimiento Grupos."}
        n_conf = int((df_g["estado_verificacion"] == "Confirmado en BD (mismo grupo)").sum())
        n_falt = int(df_g["estado_verificacion"].isin(
            ["Faltante real", "Registrado en otro grupo"]).sum())
        n_revision = int(df_g["estado_verificacion"].isin(
            ["Segundo barrido - mismo grupo", "Segundo barrido - otro grupo"]).sum())
        total = n_conf + n_falt
        pct = (n_conf / total * 100) if total else 100.0
        return {
            "grupo": grupo, "porcentaje_cumplimiento": round(pct, 1),
            "confirmados": n_conf, "faltantes": n_falt, "en_revision": n_revision,
            "cache_generado": ts,
        }

    def listar_faltantes(self, grupo: str, limite: int = 20) -> Dict[str, Any]:
        df, mensaje, ts = cargar_df_faltantes()
        if df is None:
            return {"error": mensaje}
        df_g = df[(df["grupo_original"] == grupo) & (df["estado_verificacion"] == "Faltante real")]
        if df_g.empty:
            return {"grupo": grupo, "faltantes": [],
                    "mensaje": "Sin faltantes reales registrados para este grupo."}
        filas = df_g.head(limite)
        return {
            "grupo": grupo,
            "total_faltantes": len(df_g),
            "faltantes": [
                {"producto": str(r.get("producto", ""))[:120],
                 "categoria": r.get("categoria", ""),
                 "responsable": r.get("responsable", "")}
                for _, r in filas.iterrows()
            ],
            "cache_generado": ts,
        }

    def datos_cumplimiento_todos_grupos(self) -> Dict[str, Any]:
        df, mensaje, ts = cargar_df_faltantes()
        if df is None:
            return {"error": mensaje}
        grupos_validos = [r[0] for r in self.cursor.execute('''
            SELECT DISTINCT grupo FROM grupos
            WHERE grupo IS NOT NULL AND grupo != ''
            AND grupo NOT LIKE '%SEMILLERO%' AND grupo NOT LIKE '%Semillero%'
            AND grupo NOT LIKE '%semillero%'
        ''').fetchall()]

        resultado = []
        for grupo in grupos_validos:
            df_g = df[df["grupo_original"] == grupo]
            if df_g.empty:
                continue
            n_conf = int((df_g["estado_verificacion"] == "Confirmado en BD (mismo grupo)").sum())
            n_falt = int(df_g["estado_verificacion"].isin(
                ["Faltante real", "Registrado en otro grupo"]).sum())
            total = n_conf + n_falt
            pct = (n_conf / total * 100) if total else 100.0
            resultado.append({"grupo": grupo, "pct": round(pct, 1), "faltantes": n_falt})
        resultado.sort(key=lambda r: r["pct"])
        return {"grupos": resultado, "cache_generado": ts}

    # ── Producción anual (para gráfica de líneas) ───────────────────────
    def datos_produccion_anual(self, grupo: str, anio_desde: int, anio_hasta: int) -> Dict[str, Any]:
        cedulas = [c for (c,) in self.cursor.execute(
            "SELECT DISTINCT cedula FROM grupos WHERE grupo = ?", (grupo,)).fetchall()]
        if not cedulas:
            return {"error": f"No se encontró el grupo '{grupo}'. Verifique el nombre exacto."}

        placeholders = ",".join("?" * len(cedulas))
        anios = list(range(anio_desde, anio_hasta + 1))
        idx = {a: i for i, a in enumerate(anios)}
        series = {"Publicaciones": [0] * len(anios), "Extensiones": [0] * len(anios),
                  "Proyectos": [0] * len(anios), "Trabajos de Grado": [0] * len(anios)}

        consultas = (
            ("publicaciones", "cedula", "Publicaciones"),
            ("extensiones", "cedula", "Extensiones"),
            ("proyectos", "cedula", "Proyectos"),
            ("trabajos_grado", "cedula_director", "Trabajos de Grado"),
        )
        for tabla, campo_cedula, etiqueta in consultas:
            filas = self.cursor.execute(
                f"SELECT año, COUNT(*) FROM {tabla} WHERE {campo_cedula} IN ({placeholders}) "
                f"AND año BETWEEN ? AND ? GROUP BY año",
                cedulas + [anio_desde, anio_hasta],
            ).fetchall()
            for anio, cantidad in filas:
                if anio in idx:
                    series[etiqueta][idx[anio]] = cantidad

        return {"grupo": grupo, "anios": anios, "series": series}

    def correlacion_tamano_vs_produccion(self) -> Dict[str, Any]:
        filas = self.cursor.execute('''
            SELECT grupo, COUNT(DISTINCT cedula) FROM grupos
            WHERE grupo IS NOT NULL AND grupo != ''
            AND grupo NOT LIKE '%SEMILLERO%' AND grupo NOT LIKE '%Semillero%'
            AND grupo NOT LIKE '%semillero%'
            GROUP BY grupo
        ''').fetchall()

        nombres, tamanos, producciones = [], [], []
        for grupo, tamano in filas:
            stats = self._obtener_estadisticas_grupo(grupo)
            nombres.append(grupo)
            tamanos.append(tamano)
            producciones.append(sum(stats.values()))

        if len(nombres) < 3:
            return {"error": "No hay suficientes grupos con datos para calcular una correlación."}

        r = float(np.corrcoef(tamanos, producciones)[0, 1])
        return {"grupos": nombres, "tamanos": tamanos, "producciones": producciones,
                "coeficiente_pearson": round(r, 3)}


# =============================================================================
# HERRAMIENTAS PARA EL LLM (function calling)
# =============================================================================

class HerramientasAnaliticas:
    """Define y ejecuta las herramientas que el LLM puede invocar. Cada
    ejecución devuelve (texto_json_para_el_llm, grafico_o_None,
    navegacion_o_None) -- grafico/navegacion los consume la UI, no el LLM."""

    def __init__(self, consultor: ConsultorAnalitico):
        self.consultor = consultor

    def get_tools_definition(self) -> List[Dict]:
        def tool(nombre, descripcion, propiedades, requeridos=None):
            return {"type": "function", "function": {
                "name": nombre, "description": descripcion,
                "parameters": {"type": "object", "properties": propiedades,
                                "required": requeridos or []},
            }}

        return [
            tool("listar_grupos",
                 "Lista todos los grupos de investigación con estadísticas básicas. "
                 "Úsalo cuando pregunten por grupos en general.",
                 {"filtro": {"type": "string", "description": "Texto para filtrar por nombre de grupo"}}),
            tool("detalle_grupo",
                 "Información detallada de un grupo: integrantes, publicaciones recientes, proyectos.",
                 {"nombre_grupo": {"type": "string", "description": "Nombre completo del grupo"}},
                 ["nombre_grupo"]),
            tool("analizar_produccion",
                 "Analiza producción académica interna en un rango de años (conteos, no GrupLAC).",
                 {"anio_inicio": {"type": "integer"}, "anio_fin": {"type": "integer"}},
                 ["anio_inicio", "anio_fin"]),
            tool("top_investigadores",
                 "Investigadores más productivos.",
                 {"limite": {"type": "integer", "description": "Default 10"},
                  "tipo": {"type": "string", "enum": ["todos", "publicaciones", "proyectos"]}}),
            tool("comparar_grupos",
                 "Compara dos grupos mostrando diferencias en producción.",
                 {"grupo1": {"type": "string"}, "grupo2": {"type": "string"}},
                 ["grupo1", "grupo2"]),
            tool("buscar_productos",
                 "Busca productos académicos por palabras clave en el título.",
                 {"palabras_clave": {"type": "array", "items": {"type": "string"}},
                  "tipo": {"type": "string", "enum": ["publicacion", "proyecto"]},
                  "anio_desde": {"type": "integer"}},
                 ["palabras_clave"]),
            tool("tendencias_investigacion",
                 "Tendencias de los últimos años: categorías populares y evolución temporal.",
                 {"ventana_anios": {"type": "integer", "description": "Default 5"}}),
            tool("buscar_investigador",
                 "Busca investigadores por nombre y muestra su producción.",
                 {"nombre": {"type": "string"}}, ["nombre"]),
            tool("cumplimiento_grupo",
                 "Estado de cumplimiento GrupLAC de un grupo específico (confirmados/faltantes/"
                 "en revisión), según el caché de verificación más reciente. Úsalo para preguntas "
                 "sobre si un grupo está al día en GrupLAC.",
                 {"grupo": {"type": "string"}}, ["grupo"]),
            tool("listar_faltantes",
                 "Lista los productos marcados como 'faltante real' (en BD interna, no en GrupLAC) "
                 "de un grupo, según el caché de verificación más reciente.",
                 {"grupo": {"type": "string"},
                  "limite": {"type": "integer", "description": "Default 20"}},
                 ["grupo"]),
            tool("graficar_produccion_grupo",
                 "Genera una gráfica de líneas con la producción anual de un grupo por categoría "
                 "(publicaciones, extensiones, proyectos, trabajos de grado). Úsalo cuando pidan "
                 "ver evolución/tendencia/líneas de producción de un grupo.",
                 {"grupo": {"type": "string"},
                  "anio_desde": {"type": "integer", "description": "Default: hace 5 años"},
                  "anio_hasta": {"type": "integer", "description": "Default: año actual"}},
                 ["grupo"]),
            tool("graficar_cumplimiento_grupos",
                 "Genera una gráfica de barras con el % de cumplimiento GrupLAC de TODOS los "
                 "grupos, ordenados de peor a mejor. Úsalo para preguntas generales sobre qué "
                 "grupos tienen más/menos producción confirmada o más faltantes.",
                 {}),
            tool("graficar_correlacion_tamano_produccion",
                 "Genera un diagrama de dispersión y calcula la correlación (coeficiente de "
                 "Pearson) entre el tamaño del grupo (nº integrantes) y su producción total. "
                 "Úsalo para preguntas sobre correlaciones o relaciones entre variables.",
                 {}),
            tool("cambiar_pestana",
                 "Cambia la pestaña activa de la ventana principal de la aplicación. Úsalo "
                 "cuando el usuario pida ver algo en otra parte de la app (ej. 'llévame a "
                 "Reportes por Grupo', 'muéstrame Seguimiento Grupos').",
                 {"pestana": {"type": "string",
                               "enum": ["inicio", "busqueda_personas", "reportes_grupo",
                                        "seguimiento_grupos"]},
                  "grupo": {"type": "string",
                            "description": "Grupo a preseleccionar en la pestaña destino, si aplica"}},
                 ["pestana"]),
        ]

    def ejecutar_herramienta(self, nombre: str, argumentos: Dict):
        """Devuelve (texto_json, grafico_o_None, navegacion_o_None)."""
        try:
            if nombre == "graficar_produccion_grupo":
                return self._graficar_produccion_grupo(argumentos)
            if nombre == "graficar_cumplimiento_grupos":
                return self._graficar_cumplimiento_grupos()
            if nombre == "graficar_correlacion_tamano_produccion":
                return self._graficar_correlacion()
            if nombre == "cambiar_pestana":
                navegacion = {"pestana": argumentos.get("pestana"), "grupo": argumentos.get("grupo")}
                return json.dumps({"ok": True, "accion": "cambiar_pestana"}, ensure_ascii=False), None, navegacion

            if nombre == "listar_grupos":
                resultado = self.consultor.listar_grupos(filtro_texto=argumentos.get('filtro'))
            elif nombre == "detalle_grupo":
                resultado = self.consultor.obtener_detalle_grupo(argumentos['nombre_grupo'])
            elif nombre == "analizar_produccion":
                resultado = self.consultor.analizar_produccion_por_anio(
                    argumentos['anio_inicio'], argumentos['anio_fin'])
            elif nombre == "top_investigadores":
                resultado = self.consultor.obtener_top_investigadores(
                    limite=argumentos.get('limite', 10), tipo_producto=argumentos.get('tipo', 'todos'))
            elif nombre == "comparar_grupos":
                resultado = self.consultor.comparar_grupos(argumentos['grupo1'], argumentos['grupo2'])
            elif nombre == "buscar_productos":
                resultado = self.consultor.buscar_productos_por_palabras_clave(
                    palabras=argumentos['palabras_clave'], tipo=argumentos.get('tipo'),
                    anio_desde=argumentos.get('anio_desde'))
            elif nombre == "tendencias_investigacion":
                resultado = self.consultor.obtener_tendencias_investigacion(
                    ventana_anios=argumentos.get('ventana_anios', 5))
            elif nombre == "buscar_investigador":
                resultado = self.consultor.buscar_investigador(argumentos['nombre'])
            elif nombre == "cumplimiento_grupo":
                resultado = self.consultor.cumplimiento_grupo(argumentos['grupo'])
            elif nombre == "listar_faltantes":
                resultado = self.consultor.listar_faltantes(
                    argumentos['grupo'], argumentos.get('limite', 20))
            else:
                resultado = {"error": f"Herramienta '{nombre}' no reconocida."}

            return json.dumps(resultado, ensure_ascii=False), None, None
        except KeyError as e:
            return json.dumps({"error": f"Falta el argumento requerido: {e}"}, ensure_ascii=False), None, None
        except Exception as e:
            return json.dumps({"error": str(e)}, ensure_ascii=False), None, None

    def _graficar_produccion_grupo(self, argumentos):
        anio_hasta = argumentos.get('anio_hasta') or datetime.now().year
        anio_desde = argumentos.get('anio_desde') or (anio_hasta - 5)
        datos = self.consultor.datos_produccion_anual(argumentos['grupo'], anio_desde, anio_hasta)
        if "error" in datos:
            return json.dumps(datos, ensure_ascii=False), None, None
        grafico = {"tipo": "lineas", "titulo": f"Producción anual — {datos['grupo']}",
                   "x": datos["anios"], "series": datos["series"]}
        return (json.dumps({"ok": True, "descripcion": "Gráfica de líneas generada.",
                             "datos": datos}, ensure_ascii=False),
                grafico, None)

    def _graficar_cumplimiento_grupos(self):
        datos = self.consultor.datos_cumplimiento_todos_grupos()
        if "error" in datos:
            return json.dumps(datos, ensure_ascii=False), None, None
        grafico = {"tipo": "barras_h", "titulo": "Cumplimiento GrupLAC por grupo (peor a mejor)",
                   "labels": [g["grupo"] for g in datos["grupos"]],
                   "valores": [g["pct"] for g in datos["grupos"]]}
        return (json.dumps({"ok": True, "peores_10": datos["grupos"][:10]}, ensure_ascii=False),
                grafico, None)

    def _graficar_correlacion(self):
        datos = self.consultor.correlacion_tamano_vs_produccion()
        if "error" in datos:
            return json.dumps(datos, ensure_ascii=False), None, None
        grafico = {"tipo": "dispersion", "titulo": "Tamaño del grupo vs. producción total",
                   "x": datos["tamanos"], "y": datos["producciones"], "labels": datos["grupos"],
                   "r": datos["coeficiente_pearson"]}
        return (json.dumps({"ok": True, "coeficiente_pearson": datos["coeficiente_pearson"]},
                            ensure_ascii=False),
                grafico, None)


# =============================================================================
# WORKER: conversación con Ollama (incluye el round-trip de tool calling)
# =============================================================================

class OllamaChatWorker(QThread):
    """Corre en su propio hilo: hace la(s) llamada(s) a /api/chat, y si el
    modelo pide herramientas, las ejecuta directamente acá (sqlite3 se abrió
    con check_same_thread=False) y vuelve a llamar al modelo con el
    resultado -- hasta que responda con texto final, sin más tool_calls."""

    token = pyqtSignal(str)
    done = pyqtSignal()
    error = pyqtSignal(str)
    tool_call = pyqtSignal(str, dict)
    grafico = pyqtSignal(dict)
    navegacion = pyqtSignal(dict)

    MAX_ITER_HERRAMIENTAS = 5

    def __init__(self, model, messages, herramientas: HerramientasAnaliticas,
                 usar_tools=True, host="http://localhost:11434", temperature=0.3):
        super().__init__()
        self.model = model
        self.messages = list(messages)
        self.herramientas = herramientas
        self.tools = herramientas.get_tools_definition() if usar_tools else None
        self.host = host.rstrip("/")
        self.temperature = float(temperature)

    def run(self):
        mensajes = list(self.messages)
        for _ in range(self.MAX_ITER_HERRAMIENTAS):
            resultado = self._una_llamada(mensajes)
            if resultado is None:
                return  # el error ya se emitió en _una_llamada
            contenido, tool_calls = resultado
            if not tool_calls:
                self.done.emit()
                return

            mensajes.append({"role": "assistant", "content": contenido or "", "tool_calls": tool_calls})
            for tc in tool_calls:
                nombre = tc.get("function", {}).get("name", "")
                args_raw = tc.get("function", {}).get("arguments", {})
                args = args_raw if isinstance(args_raw, dict) else json.loads(args_raw or "{}")
                self.tool_call.emit(nombre, args)

                texto, grafico, navegacion = self.herramientas.ejecutar_herramienta(nombre, args)
                if grafico is not None:
                    self.grafico.emit(grafico)
                if navegacion is not None:
                    self.navegacion.emit(navegacion)
                mensajes.append({"role": "tool", "content": texto, "name": nombre})

        self.error.emit("El modelo encadenó demasiadas herramientas sin dar una respuesta final.")

    def _una_llamada(self, mensajes):
        payload = {"model": self.model, "messages": mensajes, "stream": True,
                   "options": {"temperature": self.temperature}}
        if self.tools:
            payload["tools"] = self.tools

        try:
            r = requests.post(f"{self.host}/api/chat", json=payload, stream=True, timeout=300)
            if r.status_code == 404:
                self.error.emit(
                    f"El modelo '{self.model}' no está descargado en Ollama. "
                    f"Corra 'ollama pull {self.model}' primero.")
                return None
            r.raise_for_status()
        except requests.exceptions.ConnectionError:
            self.error.emit(
                "No se pudo conectar con Ollama en localhost:11434. "
                "¿Está corriendo el servicio? (comando: 'ollama serve')")
            return None
        except Exception as e:
            self.error.emit(str(e))
            return None

        partes: List[str] = []
        tool_calls: List[dict] = []
        for line in r.iter_lines(decode_unicode=True):
            if not line:
                continue
            try:
                data = json.loads(line)
            except (json.JSONDecodeError, TypeError):
                continue
            msg = data.get("message") or {}
            if msg.get("content"):
                partes.append(msg["content"])
                self.token.emit(msg["content"])
            if msg.get("tool_calls"):
                tool_calls.extend(msg["tool_calls"])
            if data.get("done"):
                break

        return "".join(partes), tool_calls


# =============================================================================
# VISTA: chat + panel de gráfica
# =============================================================================

class VistaChatbotInvestigacion(QWidget):
    """Pestaña de chat analítico. `cambiar_pestana_callback(pestana, grupo)`
    es inyectado desde main_10.py (VentanaPrincipal) para no acoplar este
    archivo a la ventana principal."""

    def __init__(self, db, cambiar_pestana_callback=None):
        super().__init__()
        self.db = db
        self.cambiar_pestana_callback = cambiar_pestana_callback
        self.consultor = ConsultorAnalitico(db)
        self.herramientas = HerramientasAnaliticas(self.consultor)
        self.worker: Optional[OllamaChatWorker] = None
        self.historial: List[dict] = []
        self._respuesta_actual: List[str] = []
        self._canvas_actual = None
        self.setup_ui()

    # ── UI ───────────────────────────────────────────────────────────────
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        header = QHBoxLayout()
        titulo = QLabel("<b>Asistente de Investigación</b>")
        titulo.setStyleSheet("color:#1a365d; font-size:13px;")
        header.addWidget(titulo)

        header.addWidget(QLabel("Modelo:"))
        self.combo_model = QComboBox()
        self.combo_model.setEditable(True)
        self.combo_model.addItems([MODELO_DEFAULT, "qwen2.5:3b", "llama3.1", "mistral-nemo"])
        header.addWidget(self.combo_model)

        header.addWidget(QLabel("Temp:"))
        self.spin_temp = QSpinBox()
        self.spin_temp.setRange(0, 100)
        self.spin_temp.setValue(30)
        self.spin_temp.setSuffix("%")
        self.spin_temp.setMaximumWidth(70)
        header.addWidget(self.spin_temp)

        self.btn_reset = QPushButton("Nuevo chat")
        self.btn_reset.clicked.connect(self.reset_chat)
        header.addWidget(self.btn_reset)
        header.addStretch()
        layout.addLayout(header)

        splitter = QSplitter(Qt.Horizontal)

        panel_chat = QWidget()
        lay_chat = QVBoxLayout(panel_chat)
        lay_chat.setContentsMargins(0, 0, 0, 0)
        self.chat_view = QTextEdit()
        self.chat_view.setReadOnly(True)
        self.chat_view.setStyleSheet("background:#fbfbfb; font-size:11px;")
        lay_chat.addWidget(self.chat_view, 1)

        input_layout = QHBoxLayout()
        self.input = QLineEdit()
        self.input.setPlaceholderText(
            "Pregunta sobre grupos, cumplimiento GrupLAC, producción, correlaciones...")
        self.input.returnPressed.connect(self.enviar)
        input_layout.addWidget(self.input, 1)
        self.btn_send = QPushButton("Enviar")
        self.btn_send.clicked.connect(self.enviar)
        input_layout.addWidget(self.btn_send)
        lay_chat.addLayout(input_layout)
        splitter.addWidget(panel_chat)

        self.panel_grafico = QWidget()
        self._lay_grafico = QVBoxLayout(self.panel_grafico)
        self._lbl_grafico_vacio = QLabel(
            "Las gráficas que pidas aparecerán acá\n"
            "(ej. \"muéstrame el cumplimiento de todos los grupos\",\n"
            "\"grafica la producción del grupo X\",\n"
            "\"hay correlación entre tamaño y producción?\").")
        self._lbl_grafico_vacio.setAlignment(Qt.AlignCenter)
        self._lbl_grafico_vacio.setStyleSheet("color:#898781; font-size:11px; padding:20px;")
        self._lay_grafico.addWidget(self._lbl_grafico_vacio)
        splitter.addWidget(self.panel_grafico)

        splitter.setSizes([600, 500])
        layout.addWidget(splitter, 1)

        self._append_sistema(
            "Hola. Soy el asistente de análisis de investigación. Puedo consultar la base de "
            "datos interna y el estado de cumplimiento GrupLAC más reciente, generar gráficas "
            "(producción por grupo, cumplimiento de todos los grupos, correlación tamaño/"
            "producción) y llevarte a otra pestaña de la app si lo pides.\n\n"
            "Necesita Ollama corriendo localmente con un modelo que soporte herramientas "
            f"(recomendado: {MODELO_DEFAULT})."
        )

    def _append_sistema(self, texto):
        self.chat_view.append(
            f"<div style='background:#e3f2fd; padding:8px; margin:4px; border-radius:5px; "
            f"border-left:4px solid #2196f3;'>{texto}</div>")

    def _append_usuario(self, texto):
        self.chat_view.append(
            f"<div style='background:#f5f5f5; padding:8px; margin:4px; border-radius:5px; "
            f"border-left:4px solid #4caf50;'><b>Tú:</b> {texto}</div>")

    def _append_inicio_asistente(self):
        self.chat_view.append(
            "<div style='background:#fff3e0; padding:8px; margin:4px; border-radius:5px; "
            "border-left:4px solid #ff9800;'><b>Asistente:</b> ")

    def _append_herramienta(self, nombre, args):
        args_str = json.dumps(args, ensure_ascii=False)
        self.chat_view.append(
            f"<div style='font-size:9px; color:#666; margin:2px 4px;'>"
            f"Consultando: <code>{nombre}</code> {args_str}</div>")

    def _append_error(self, mensaje):
        self.chat_view.append(
            f"<div style='background:#ffebee; padding:8px; margin:4px; border-radius:5px; "
            f"border-left:4px solid #f44336;'><b>Error:</b> {mensaje}</div>")

    # ── Conversación ─────────────────────────────────────────────────────
    def reset_chat(self):
        self.chat_view.clear()
        self.historial = []
        self._append_sistema("Conversación reiniciada. ¿En qué puedo ayudarte?")

    def enviar(self):
        texto = self.input.text().strip()
        if not texto or self.worker is not None:
            return

        self.input.clear()
        self._append_usuario(texto)
        self._append_inicio_asistente()
        self._respuesta_actual = []

        mensajes = self._construir_mensajes(texto)
        self.historial.append({"role": "user", "content": texto})

        self.btn_send.setEnabled(False)
        self.input.setEnabled(False)

        modelo = self.combo_model.currentText().strip() or MODELO_DEFAULT
        temperatura = self.spin_temp.value() / 100.0

        self.worker = OllamaChatWorker(
            model=modelo, messages=mensajes, herramientas=self.herramientas,
            temperature=temperatura)
        self.worker.token.connect(self._on_token)
        self.worker.tool_call.connect(self._append_herramienta)
        self.worker.grafico.connect(self._renderizar_grafico)
        self.worker.navegacion.connect(self._on_navegacion)
        self.worker.done.connect(self._on_done)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def _construir_mensajes(self, texto_usuario):
        system = (
            "Eres un asistente analítico especializado en la producción académica e "
            "investigativa de la Universidad Tecnológica de Pereira.\n"
            "- Usa las herramientas disponibles para consultar datos reales antes de responder; "
            "no inventes cifras.\n"
            "- Para preguntas sobre cumplimiento, faltantes o si algo está subido a GrupLAC, usa "
            "las herramientas de GrupLAC (cumplimiento_grupo, listar_faltantes), no supongas.\n"
            "- Si piden ver una gráfica, tendencia, evolución o correlación, usa la herramienta "
            "graficar_* correspondiente en vez de solo describir números en texto. Esa "
            "herramienta YA dibuja la gráfica en un panel aparte de la pantalla -- nunca "
            "inventes ni escribas un link o markdown de imagen (![...](...)) en tu respuesta, "
            "la gráfica no se muestra dentro del texto.\n"
            "- Si piden ir a otra parte de la aplicación (Reportes por Grupo, Seguimiento "
            "Grupos, Búsqueda de Personas, Inicio), usa cambiar_pestana.\n"
            "- Responde en español, conciso, con **negritas** para lo importante."
        )
        mensajes = [{"role": "system", "content": system}]
        mensajes.extend(self.historial[-12:])
        mensajes.append({"role": "user", "content": texto_usuario})
        return mensajes

    def _on_token(self, chunk):
        self._respuesta_actual.append(chunk)
        cursor = self.chat_view.textCursor()
        cursor.movePosition(cursor.End)
        cursor.insertText(chunk)
        self.chat_view.setTextCursor(cursor)
        self.chat_view.ensureCursorVisible()

    def _on_navegacion(self, navegacion):
        if self.cambiar_pestana_callback is None:
            return
        self.cambiar_pestana_callback(navegacion.get("pestana"), navegacion.get("grupo"))

    def _on_done(self):
        self.chat_view.append("</div>")
        texto_final = "".join(self._respuesta_actual).strip()
        if texto_final:
            self.historial.append({"role": "assistant", "content": texto_final})
        self.worker = None
        self.btn_send.setEnabled(True)
        self.input.setEnabled(True)
        self.input.setFocus()

    def _on_error(self, mensaje):
        self.chat_view.append("</div>")
        self._append_error(mensaje)
        self.worker = None
        self.btn_send.setEnabled(True)
        self.input.setEnabled(True)

    # ── Gráficas ─────────────────────────────────────────────────────────
    def _renderizar_grafico(self, datos: dict):
        from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
        from matplotlib.figure import Figure

        if self._canvas_actual is not None:
            self._lay_grafico.removeWidget(self._canvas_actual)
            self._canvas_actual.setParent(None)
        self._lbl_grafico_vacio.hide()

        fig = Figure(figsize=(5.5, 4.2), facecolor="#fcfcfb")
        ax = fig.add_subplot(111)
        ax.set_facecolor("#fcfcfb")
        tipo = datos.get("tipo")

        if tipo == "lineas":
            for etiqueta, valores in datos.get("series", {}).items():
                ax.plot(datos.get("x", []), valores, marker="o", label=etiqueta, linewidth=1.5)
            ax.legend(fontsize=7)
            ax.set_xlabel("Año", fontsize=8)
        elif tipo == "barras_h":
            labels = datos.get("labels", [])
            valores = datos.get("valores", [])
            n = len(labels)
            y = list(range(n))
            colores = [self._color_para_pct(v) for v in valores]
            ax.barh(y, valores, color=colores, height=0.7)
            ax.set_yticks(y)
            ax.set_yticklabels(labels, fontsize=5 if n > 40 else 7)
            ax.invert_yaxis()
            ax.set_xlim(0, 100)
            ax.set_xlabel("% Cumplimiento", fontsize=8)
        elif tipo == "dispersion":
            ax.scatter(datos.get("x", []), datos.get("y", []), color="#2a78d6", alpha=0.75)
            ax.set_xlabel("Tamaño del grupo (integrantes)", fontsize=8)
            ax.set_ylabel("Producción total", fontsize=8)
            if "r" in datos:
                ax.text(0.05, 0.93, f"r = {datos['r']:.2f}", transform=ax.transAxes,
                        va="top", fontsize=9, color="#1a365d")

        ax.set_title(datos.get("titulo", ""), fontsize=9, color="#1a365d")
        ax.tick_params(axis="both", labelsize=7, colors="#52514e")
        for spine in ("top", "right"):
            ax.spines[spine].set_visible(False)
        fig.tight_layout()

        self._canvas_actual = FigureCanvas(fig)
        self._lay_grafico.addWidget(self._canvas_actual)

    @staticmethod
    def _color_para_pct(pct):
        if pct >= 90:
            return "#0ca30c"
        if pct >= 70:
            return "#fab219"
        if pct >= 50:
            return "#ec835a"
        return "#d03b3b"
