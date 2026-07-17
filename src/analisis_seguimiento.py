"""
Análisis de seguimiento de grupos de investigación.

Módulos:
  - AnalisisDuplicados: detecta autores/productos duplicados entre grupos.
  - CategoriaAnalyzer957: analiza posición de un grupo según indicadores de
      la convocatoria 957, usando datame/gruplac_957.db como referencia.
  - SimuladorCategoriaInterna: simula la categoría que tendría un grupo con
      base en los productos registrados en la BD interna y las ventanas de
      evaluación oficiales de MinCiencias.
"""

import json
import math
import re
import sqlite3
from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd
from unidecode import unidecode

from constants import (
    VENTANAS_957,
    INDICADORES_957,
    ORDEN_CATEGORIAS_MINCIENCIAS,
    CODIGO_957_A_INDICADOR,
    INDICADOR_957_A_CUARTIL,
    CUARTIL_OBJETIVO_POR_CATEGORIA,
    SECCION_957_POR_INDICADOR,
)
from utils import obtener_directorio_base


# =============================================================================
# UTILIDADES LOCALES
# =============================================================================

def _normalizar(texto: str) -> str:
    if not texto:
        return ""
    return re.sub(r"\s+", " ", unidecode(str(texto).lower())).strip(" ,.;:!?-_")


def _similitud_jaccard(t1: str, t2: str) -> float:
    """Similitud de Jaccard sobre tokens normalizados."""
    n1, n2 = _normalizar(t1), _normalizar(t2)
    if not n1 or not n2:
        return 0.0
    s1, s2 = set(n1.split()), set(n2.split())
    if not s1 or not s2:
        return 0.0
    return len(s1 & s2) / len(s1 | s2)


# =============================================================================
# DETECCIÓN DE DUPLICADOS
# =============================================================================

class AnalisisDuplicados:
    """Detecta autores y productos duplicados en la base de datos interna."""

    TIPOS_RETIRADO = {"ex integrante", "exintegrante", "retirado", "retirada"}

    def __init__(self, db_conn):
        self.conn = db_conn

    def autores_duplicados(self) -> pd.DataFrame:
        """
        Devuelve personas que aparecen en más de un grupo.
        Columnas: cedula, nombre, grupos, num_grupos
        Excluye miembros con tipo_miembro en TIPOS_RETIRADO.
        """
        retirados = "','".join(self.TIPOS_RETIRADO)
        query = f"""
            SELECT g.cedula, p.nombre, g.grupo, g.facultad
            FROM grupos g
            JOIN personas p ON g.cedula = p.cedula
            WHERE g.grupo IS NOT NULL AND g.grupo != ''
              AND (g.tipo_miembro IS NULL OR g.tipo_miembro NOT IN ('{retirados}'))
        """
        df = pd.read_sql_query(query, self.conn)
        if df.empty:
            return pd.DataFrame(columns=["cedula", "nombre", "grupos", "num_grupos"])

        agrupado = (
            df.groupby(["cedula", "nombre"])
            .agg(grupo=("grupo", list), facultad=("facultad", list))
            .reset_index()
        )
        agrupado["num_grupos"] = agrupado["grupo"].apply(len)
        agrupado = agrupado[agrupado["num_grupos"] > 1].copy()
        agrupado["grupos"] = agrupado["grupo"].apply(lambda lst: "; ".join(lst))
        return (
            agrupado[["cedula", "nombre", "grupos", "num_grupos"]]
            .sort_values("num_grupos", ascending=False)
        )

    # ── Redistribución de productos entre grupos ─────────────────────────

    @staticmethod
    def _normalizar_grupo(nombre: str) -> str:
        """Normaliza nombre de grupo para matching cross-origen."""
        return re.sub(r'\s+', ' ', unidecode(str(nombre).lower().strip()))

    def _cargar_integrantes_gruplac(self, ruta_base=None):
        """
        Lee hoja 'Integrantes del grupo' de cada xlsx en reporte excel/.
        Retorna: dict[grupo_normalizado] → list of {
            nombre_original, nombre_norm, vinculacion, horas, inicio_fin,
            desde, hasta, activo
        }
        """
        import pandas as _pd
        from pathlib import Path as _Path

        if ruta_base is None:
            from utils import obtener_directorio_base as _odb
            ruta_base = _odb() / "reports" / "excel"

        ruta = _Path(ruta_base)
        if not ruta.exists():
            self._integrantes_gruplac = {}
            return {}

        resultado = {}
        for xlsx in sorted(ruta.glob("**/*.xlsx")):
            grupo = self._normalizar_grupo(xlsx.parent.name)
            try:
                xl = _pd.ExcelFile(xlsx, engine='openpyxl')
                if "Integrantes del grupo" not in xl.sheet_names:
                    continue
                df = xl.parse("Integrantes del grupo")
                if df.empty:
                    continue

                # Identificar columnas por patrón
                fecha_col = None
                vinc_col = None
                for col in df.columns:
                    col_s = str(col).encode('latin-1', errors='replace').decode('latin-1')
                    if 'inicio' in col_s.lower():
                        fecha_col = col
                    if 'vinculaci' in col_s.lower():
                        vinc_col = col

                integrantes = []
                for _, row in df.iterrows():
                    nombre_raw = str(row.get(df.columns[0], ""))
                    if not nombre_raw or nombre_raw.lower() in ("", "nan", "none", "nombres"):
                        continue
                    nombre_limpio = nombre_raw.replace("\n", " ").strip()
                    nombre_limpio = re.sub(r"^\d+\s*\.?-?\s*", "", nombre_limpio).strip()
                    nombre_norm = _normalizar(nombre_limpio)
                    if not nombre_norm:
                        continue

                    entry = {
                        "nombre_original": nombre_limpio,
                        "nombre_norm": nombre_norm,
                        "vinculacion": str(row.get(vinc_col, "")) if vinc_col else "",
                        "horas": "",
                        "inicio_fin": "",
                        "desde": "—",
                        "hasta": "—",
                        "activo": False,
                    }

                    # Parsear fecha
                    if fecha_col:
                        raw = str(row.get(fecha_col, ""))
                        entry["inicio_fin"] = raw
                        if " - " in raw:
                            partes = raw.split(" - ", 1)
                            entry["desde"] = partes[0].strip()
                            entry["hasta"] = partes[1].strip()
                            entry["activo"] = "actual" in partes[1].lower()

                    integrantes.append(entry)

                if integrantes:
                    resultado[grupo] = integrantes

            except Exception:
                continue

        self._integrantes_gruplac = resultado
        return resultado

    def _obtener_cache_integrantes(self):
        cache_integrantes = getattr(self, "_integrantes_gruplac", None)
        if not cache_integrantes:
            cache_integrantes = self._cargar_integrantes_gruplac()
        return cache_integrantes

    @staticmethod
    def _nivel_coincidencia_nombre(nombre_norm: str, entry_nombre_norm: str) -> int:
        """
        Nivel de confianza de la coincidencia entre dos nombres normalizados:
          3 = exacto.
          2 = apellido + primera inicial del nombre (confiable).
          1 = 2 o más apellidos/nombres en común (>3 caracteres) — heurística
              que genera falsos positivos con apellidos comunes, requiere
              verificación manual.
          0 = no coincide.
        """
        if nombre_norm == entry_nombre_norm:
            return 3

        partes_q = nombre_norm.split()
        partes_e = entry_nombre_norm.split()
        if len(partes_q) >= 2 and len(partes_e) >= 2:
            if partes_e[-1] == partes_q[-1] and partes_e[0][:1] == partes_q[0][:1]:
                return 2

        apellidos_q = {p for p in partes_q if len(p) > 3 and p not in (
            "de", "la", "del", "los", "las", "san", "santa")}
        if len(apellidos_q) >= 2 and len(apellidos_q & set(partes_e)) >= 2:
            return 1

        return 0

    @classmethod
    def _coincide_nombre(cls, nombre_norm: str, entry_nombre_norm: str, estricto: bool = False) -> bool:
        """
        `estricto=True` exige nivel >= 2 (exacto o apellido+inicial): razonable
        al buscar en TODO GrupLAC (miles de personas), donde dos apellidos
        comunes coinciden por azar con facilidad. `estricto=False` también
        acepta el nivel 1 (heurístico), pensado para buscar dentro de UN grupo
        (decenas de personas) — ver `_nombre_en_gruplac`, que marca ese nivel
        como "probable" en vez de "confirmado" para que se verifique a mano.
        """
        nivel = cls._nivel_coincidencia_nombre(nombre_norm, entry_nombre_norm)
        return nivel >= (2 if estricto else 1)

    def resolver_grupo_en_gruplac(self, grupo, cache_integrantes=None):
        """
        Encuentra la clave de `cache_integrantes` que corresponde a `grupo`,
        tolerando que el nombre de carpeta GrupLAC venga truncado (algunos
        reportes se generaron con nombres de grupo muy largos y el sistema de
        archivos cortó el nombre de la carpeta). Devuelve la clave encontrada
        o None si no hay ningún reporte GrupLAC para ese grupo.
        """
        cache_integrantes = cache_integrantes or self._obtener_cache_integrantes()
        grupo_norm = self._normalizar_grupo(grupo)
        if grupo_norm in cache_integrantes:
            return grupo_norm

        # Solo se intenta este "fallback" con nombres largos: los truncamientos
        # de carpeta ocurren por límite de longitud de ruta, no en nombres cortos.
        if len(grupo_norm) < 20:
            return None
        candidatos = [
            k for k in cache_integrantes
            if len(k) >= 20 and (k.startswith(grupo_norm) or grupo_norm.startswith(k))
        ]
        if not candidatos:
            return None
        return max(candidatos, key=len)

    def _nombre_en_gruplac(self, nombre_norm, grupo, cache_integrantes=None):
        """
        Busca una persona (nombre normalizado) en los integrantes GrupLAC de un grupo.
        Retorna dict con info o None si no se encuentra. Ver `resolver_grupo_en_gruplac`
        para la tolerancia a nombres de carpeta truncados.

        Solo acepta nivel >= 2 (exacto o apellido+inicial): el nivel 1
        (2+ apellidos en común) no es confiable para afirmar presencia —
        p.ej. "Mauricio Holguín Londoño" y "Germán Andrés Holguín Londoño"
        comparten ambos apellidos pero son personas distintas. El software no
        debe sugerir una coincidencia dudosa con la persona equivocada; si no
        hay coincidencia confiable, se reporta como "no aparece".
        """
        cache_integrantes = cache_integrantes or self._obtener_cache_integrantes()
        grupo_norm = self.resolver_grupo_en_gruplac(grupo, cache_integrantes)
        if grupo_norm is None:
            return None
        for entry in cache_integrantes.get(grupo_norm, []):
            if self._coincide_nombre(nombre_norm, entry["nombre_norm"], estricto=True):
                return entry
        return None

    def buscar_persona_en_gruplac(self, nombre_norm, cache_integrantes=None,
                                   excluir_grupos_norm=None):
        """
        Busca una persona (nombre normalizado) en TODOS los grupos GrupLAC
        cargados (no solo uno). Útil para detectar a alguien que aparece en
        GrupLAC en un grupo donde no está registrado en la BD interna.

        Exige coincidencia EXACTA (nivel 3): comparado contra miles de
        personas de todos los grupos a la vez, incluso "apellido + inicial"
        (nivel 2) produce falsos positivos (cualquier "M. ... Londoño" de
        otro grupo dispara una coincidencia que no es la misma persona).
        Devuelve list[(grupo_norm, entry)], una entrada por grupo donde aparece.
        """
        cache_integrantes = cache_integrantes or self._obtener_cache_integrantes()
        excluir = excluir_grupos_norm or set()
        hallazgos = []
        for grupo_norm, integrantes in cache_integrantes.items():
            if grupo_norm in excluir:
                continue
            for entry in integrantes:
                if self._nivel_coincidencia_nombre(nombre_norm, entry["nombre_norm"]) >= 3:
                    hallazgos.append((grupo_norm, entry))
                    break
        return hallazgos

    def personas_gruplac_sin_registro_interno(self, ruta_gruplac=None,
                                               solo_activos: bool = True) -> pd.DataFrame:
        """
        Dirección inversa a `_nombre_en_gruplac`: para cada grupo con reporte
        GrupLAC, lista los integrantes que GrupLAC SÍ tiene pero que no
        calzan con nadie de la tabla interna `grupos` para ese mismo grupo
        (nivel >= 2: exacto o apellido+inicial — mismo umbral ya validado
        para no adivinar coincidencias dudosas). Solo reporta; no agrega
        nada a la BD interna, la decisión de qué hacer queda para quien
        revise el reporte.

        `solo_activos=True` (default) limita a integrantes con vinculación
        vigente ("... - Actual") en GrupLAC -- la hoja "Integrantes del
        grupo" trae TODO el historial (miles de personas que ya se fueron
        hace años), y comparar contra eso genera ruido masivo de gente que
        con toda razón no está en la BD interna actual.

        Columnas: grupo, nombre_gruplac, vinculacion.
        """
        cache_integrantes = self._cargar_integrantes_gruplac(ruta_gruplac)
        cursor = self.conn.cursor()
        grupos_internos = [
            g for (g,) in cursor.execute(
                "SELECT DISTINCT grupo FROM grupos WHERE grupo IS NOT NULL AND grupo != ''")
        ]

        filas = []
        for grupo in grupos_internos:
            grupo_norm = self.resolver_grupo_en_gruplac(grupo, cache_integrantes)
            if grupo_norm is None:
                continue  # sin reporte GrupLAC para este grupo -> no se puede comparar
            integrantes_gl = cache_integrantes.get(grupo_norm, [])
            if solo_activos:
                integrantes_gl = [e for e in integrantes_gl if e.get("activo")]
            if not integrantes_gl:
                continue
            nombres_internos = [
                _normalizar(nombre) for (nombre,) in cursor.execute(
                    "SELECT DISTINCT p.nombre FROM grupos g "
                    "JOIN personas p ON g.cedula = p.cedula WHERE g.grupo = ?", (grupo,))
            ]
            for entry in integrantes_gl:
                existe = any(
                    self._coincide_nombre(entry["nombre_norm"], nombre_interno, estricto=True)
                    for nombre_interno in nombres_internos
                )
                if not existe:
                    filas.append({
                        "grupo": grupo,
                        "nombre_gruplac": entry["nombre_original"],
                        "vinculacion": entry.get("vinculacion", ""),
                    })

        return pd.DataFrame(filas, columns=["grupo", "nombre_gruplac", "vinculacion"])

    def _productos_de_persona(self, cedula: str) -> List[Dict]:
        """Todos los productos de una persona en la BD interna."""
        prods = []

        # (tabla, campo_id, columna_titulo, columna_tipo)
        _CONFIG = [
            ("publicaciones",        "cedula",           "titulo",      "tipo"),
            ("extensiones",          "cedula",           "actividad",   "tipo"),
            ("trabajos_grado",       "cedula_director",  "titulo",      "tipo"),
            ("proyectos",            "cedula",           "titulo",      "tipo"),
            ("productos_innovacion", "cedula",           "nombre",      "tipo_producto"),
        ]

        for tabla, campo_id, col_tit, col_tip in _CONFIG:
            try:
                cursor = self.conn.cursor()
                rows = cursor.execute(
                    f"SELECT {col_tit}, año, {col_tip}, grupo FROM {tabla} "
                    f"WHERE {campo_id} = ? AND {col_tit} IS NOT NULL AND {col_tit} != ''",
                    (cedula,)
                ).fetchall()
                for r in rows:
                    prods.append({
                        "titulo": r[0],
                        "anio": r[1],
                        "tipo": r[2] or "",
                        "grupo": r[3] or "",
                        "fuente": tabla,
                    })
            except Exception:
                pass

        # También buscar como estudiante en trabajos_grado
        try:
            cursor = self.conn.cursor()
            rows = cursor.execute(
                "SELECT titulo, año, tipo, grupo FROM trabajos_grado "
                "WHERE cedula_estudiante = ? AND titulo IS NOT NULL AND titulo != ''",
                (cedula,)
            ).fetchall()
            for r in rows:
                prods.append({
                    "titulo": r[0],
                    "anio": r[1],
                    "tipo": r[2] or "",
                    "grupo": r[3] or "",
                    "fuente": "trabajos_grado",
                })
        except Exception:
            pass

        # Clasificar tipo_957 desde VENTANAS_957
        from constants import VENTANAS_957 as _V
        for p in prods:
            tipo_orig = p.get("tipo", "")
            p["tipo_957"] = ""
            p["lambda_957"] = 0
            p["indicador_957"] = ""
            for key in _V:
                if key.lower() in tipo_orig.lower():
                    p["tipo_957"] = key
                    p["lambda_957"] = _V[key]["lambda"]
                    p["indicador_957"] = _V[key]["indicador"]
                    break

        return prods

    @staticmethod
    def _producto_en_gruplac(titulo_norm: str, grupo: str,
                              datos_gruplac: Dict) -> bool:
        """Verifica si un título normalizado existe en los detalles GrupLAC
        de un grupo (matching exacto de título normalizado)."""
        detalles = datos_gruplac.get(grupo, {}).get("detalles", [])
        for d in detalles:
            d_norm = _normalizar(d.get("titulo", ""))
            if d_norm == titulo_norm:
                return True
        return False

    def _productos_persona_por_grupo(self, cedula, grupos_persona, datos_gruplac):
        """
        Para cada producto de una persona, indica en qué grupos GrupLAC
        está reclamado (match exacto de título normalizado).
        Retorna list[dict]: producto, anio, tipo_957, lambda, indicador,
        reclamado_en, no_reclamado_en
        """
        prods = self._productos_de_persona(cedula)
        resultado = []
        for prod in prods:
            t_norm = _normalizar(prod["titulo"])
            if len(t_norm) < 10:
                continue
            reclamado_en = []
            no_reclamado_en = []
            for g in grupos_persona:
                if self._producto_en_gruplac(t_norm, g, datos_gruplac):
                    reclamado_en.append(g)
                else:
                    no_reclamado_en.append(g)
            resultado.append({
                "producto": prod["titulo"],
                "anio": prod.get("anio", ""),
                "tipo": prod.get("tipo", ""),
                "tipo_957": prod.get("tipo_957", ""),
                "lambda": prod.get("lambda_957", 0),
                "indicador": prod.get("indicador_957", ""),
                "reclamado_en": reclamado_en,
                "no_reclamado_en": no_reclamado_en,
            })
        return resultado

    def oportunidades_redistribucion(
        self,
        datos_gruplac: Dict,
        productos_clasificados: Dict,
        anio_base: Optional[int] = None,
    ) -> pd.DataFrame:
        """
        Para cada persona en múltiples grupos, detecta productos BD que están
        en el GrupLAC de un grupo pero no en el de otro → oportunidad de
        reclamación.

        Returns DataFrame con columnas:
        miembro, cedula, producto, tipo_957, lambda, indicador, anio,
        grupo_origen (lista), grupo_destino, en_ventana, fuente
        """
        if anio_base is None:
            anio_base = datetime.now().year
        df_multi = self.autores_duplicados()
        if df_multi.empty:
            return pd.DataFrame(columns=[
                "miembro", "cedula", "producto", "tipo_957",
                "lambda", "indicador", "anio", "grupo_origen",
                "grupo_destino", "en_ventana", "fuente",
            ])

        from collections import defaultdict
        from constants import VENTANAS_957 as _V

        filas = []
        for _, row in df_multi.iterrows():
            cedula = row["cedula"]
            grupos_persona = [g.strip() for g in row["grupos"].split("; ")]

            for prod in self._productos_de_persona(cedula):
                t_norm = _normalizar(prod["titulo"])
                if len(t_norm) < 10:
                    continue

                # Determinar en qué grupos GrupLAC está este producto
                grupos_con = []
                for g in grupos_persona:
                    if self._producto_en_gruplac(t_norm, g, datos_gruplac):
                        grupos_con.append(g)

                if not grupos_con or len(grupos_con) == len(grupos_persona):
                    continue  # sin oportunidad

                grupos_sin = [g for g in grupos_persona if g not in grupos_con]

                tipo_957 = prod.get("tipo", "")
                cfg = _V.get(tipo_957)
                if not cfg:
                    continue

                anio = prod.get("anio")
                en_ventana = False
                try:
                    anio_int = int(anio)
                    en_ventana = (
                        anio_base - cfg["ventana"] < anio_int <= anio_base
                    )
                except (ValueError, TypeError):
                    pass

                for g_dest in grupos_sin:
                    filas.append({
                        "miembro": row["nombre"],
                        "cedula": cedula,
                        "producto": prod["titulo"][:100],
                        "tipo_957": tipo_957,
                        "lambda": cfg["lambda"],
                        "indicador": cfg["indicador"],
                        "anio": anio,
                        "grupo_origen": "; ".join(sorted(grupos_con)),
                        "grupo_destino": g_dest,
                        "en_ventana": en_ventana,
                        "fuente": prod["fuente"],
                    })

        if not filas:
            return pd.DataFrame(columns=[
                "miembro", "cedula", "producto", "tipo_957",
                "lambda", "indicador", "anio", "grupo_origen",
                "grupo_destino", "en_ventana", "fuente",
            ])

        return pd.DataFrame(filas).sort_values(
            ["grupo_destino", "indicador", "lambda"],
            ascending=[True, True, False],
        ).reset_index(drop=True)

    def productos_duplicados(self, umbral_similitud: float = 0.75) -> pd.DataFrame:
        """
        Detecta productos que aparecen en múltiples grupos comparando:
        1) DOI exacto  2) ISSN/ISBN exacto  3) Similitud de título ≥ umbral

        Devuelve DataFrame con columnas:
        titulo, año, grupos, autores, similitud_max, match_tipo, count
        """
        tablas = [
            "publicaciones", "extensiones", "trabajos_grado",
            "productos_innovacion", "proyectos",
        ]
        dfs = []
        for tabla in tablas:
            df_tabla = self._leer_tabla(tabla)
            if df_tabla is not None and not df_tabla.empty:
                dfs.append(df_tabla)

        if not dfs:
            return pd.DataFrame(
                columns=["titulo", "año", "grupos", "autores",
                         "similitud_max", "match_tipo", "count"]
            )

        todos = pd.concat(dfs, ignore_index=True)
        todos["titulo_norm"] = todos["titulo"].apply(_normalizar)
        todos = todos[todos["titulo_norm"].str.len() >= 10]

        if todos.empty:
            return pd.DataFrame(
                columns=["titulo", "año", "grupos", "autores",
                         "similitud_max", "match_tipo", "count"]
            )

        resultados = []
        resultados.extend(self._duplicados_por_campo(todos, "doi_url", "DOI"))
        resultados.extend(self._duplicados_por_campo(todos, "issn_isbn", "ISSN/ISBN"))
        resultados.extend(
            self._duplicados_por_titulo(todos, umbral_similitud)
        )

        if not resultados:
            return pd.DataFrame(
                columns=["titulo", "año", "grupos", "autores",
                         "similitud_max", "match_tipo", "count"]
            )

        resultado_df = pd.concat(resultados, ignore_index=True, sort=False)

        def unir(lst):
            return "; ".join(sorted(set(str(x) for x in lst if pd.notna(x))))

        resultado_df["titulo"] = resultado_df["titulo"].apply(
            lambda lst: lst[0] if isinstance(lst, list) and lst else str(lst)
        )
        resultado_df["año"] = resultado_df["año"].apply(unir)
        resultado_df["grupos"] = resultado_df["grupo"].apply(unir)
        resultado_df["autores"] = resultado_df["cedula"].apply(unir)
        # Resolver cédulas a nombres reales desde la tabla personas
        try:
            cursor = self.conn.cursor()
            todas = set()
            for val in resultado_df["autores"]:
                for c in str(val).split("; "):
                    c = c.strip()
                    if c:
                        todas.add(c)
            if todas:
                placeholders = ",".join("?" for _ in todas)
                filas = cursor.execute(
                    f"SELECT cedula, nombre FROM personas WHERE cedula IN ({placeholders})",
                    list(todas)
                ).fetchall()
                mapa = dict(filas)
                def _reemplazar(val):
                    return "; ".join(
                        mapa.get(c.strip(), c.strip())
                        for c in str(val).split("; ") if c.strip()
                    )
                resultado_df["autores"] = resultado_df["autores"].apply(_reemplazar)
        except Exception:
            pass
        resultado_df = resultado_df.rename(columns={"match_por": "match_tipo"})

        cols = ["titulo", "año", "grupos", "autores", "similitud_max", "match_tipo", "count"]
        for col in cols:
            if col not in resultado_df.columns:
                resultado_df[col] = None
        return resultado_df[cols].sort_values(["match_tipo", "count"], ascending=[True, False])

    # ── helpers privados ─────────────────────────────────────────────────────

    @staticmethod
    def _col_titulo(tabla: str) -> str:
        return {
            "extensiones": "actividad",
            "productos_innovacion": "nombre",
        }.get(tabla, "titulo")

    def _leer_tabla(self, tabla: str) -> Optional[pd.DataFrame]:
        """Intenta leer la tabla con las columnas esperadas."""
        col_tit = self._col_titulo(tabla)

        if tabla == "trabajos_grado":
            query = f"""
                SELECT '{tabla}' AS fuente,
                       cedula_director AS cedula,
                       titulo,
                       año,
                       grupo,
                       '' AS autores,
                       '' AS doi_url,
                       '' AS issn_isbn
                FROM {tabla}
                WHERE grupo IS NOT NULL AND grupo != '' AND titulo IS NOT NULL AND titulo != ''
                UNION
                SELECT '{tabla}' AS fuente,
                       cedula_estudiante AS cedula,
                       titulo,
                       año,
                       grupo,
                       '' AS autores,
                       '' AS doi_url,
                       '' AS issn_isbn
                FROM {tabla}
                WHERE grupo IS NOT NULL AND grupo != '' AND titulo IS NOT NULL AND titulo != ''
            """
        else:
            query = f"""
                SELECT '{tabla}' AS fuente,
                       cedula,
                       {col_tit} AS titulo,
                       año,
                       grupo,
                       '' AS autores,
                       '' AS doi_url,
                       '' AS issn_isbn
                FROM {tabla}
                WHERE grupo IS NOT NULL AND grupo != '' AND {col_tit} IS NOT NULL AND {col_tit} != ''
            """
        try:
            return pd.read_sql_query(query, self.conn)
        except Exception:
            pass

        try:
            df_raw = pd.read_sql_query(
                f"SELECT * FROM {tabla} WHERE grupo IS NOT NULL AND grupo != ''",
                self.conn,
            )
            df = pd.DataFrame()
            df["fuente"] = tabla

            if tabla == "trabajos_grado":
                cedula_dr = df_raw.get("cedula_director", pd.Series())
                cedula_st = df_raw.get("cedula_estudiante", pd.Series())
                df["cedula"] = cedula_dr if not cedula_dr.isna().all() else cedula_st
            else:
                df["cedula"] = df_raw.get("cedula", df_raw.get("Cedula", pd.Series()))

            titulo_raw = df_raw.get(col_tit, None)
            if titulo_raw is None:
                titulo_raw = df_raw.get(col_tit.capitalize(), None)
            if titulo_raw is None and tabla == "extensiones":
                titulo_raw = df_raw.get("Actividad", None)
            if titulo_raw is None and tabla == "productos_innovacion":
                titulo_raw = df_raw.get("Nombre", None)
            if titulo_raw is None:
                titulo_raw = df_raw.get("titulo", df_raw.get("Titulo", pd.Series()))
            df["titulo"] = titulo_raw

            df["año"] = df_raw.get("año", df_raw.get("Año", df_raw.get("anio", pd.Series())))
            df["grupo"] = df_raw.get("grupo", df_raw.get("Grupo", pd.Series()))
            df["autores"] = ""
            df["doi_url"] = None
            df["issn_isbn"] = None
            return df.dropna(subset=["titulo", "grupo"])
        except Exception:
            return None

    @staticmethod
    def _duplicados_por_campo(
        df: pd.DataFrame, campo: str, etiqueta: str
    ) -> List[pd.DataFrame]:
        if campo not in df.columns:
            return []
        sub = df[df[campo].notna() & (df[campo] != "")]
        if sub.empty:
            return []
        grp = (
            sub.groupby(campo)
            .agg(titulo=("titulo", list), grupo=("grupo", list),
                 cedula=("cedula", list), año=("año", list))
            .reset_index()
        )
        grp["count"] = grp["grupo"].apply(len)
        grp = grp[grp["count"] > 1].copy()
        if grp.empty:
            return []
        grp["match_por"] = etiqueta
        grp["similitud_max"] = 1.0
        return [grp]

    @staticmethod
    def _duplicados_por_titulo(
        df: pd.DataFrame, umbral: float
    ) -> List[pd.DataFrame]:
        """
        Detecta títulos duplicados usando índice invertido de tokens.
        Solo hace comparación completa entre pares que comparten ≥1 token,
        evitando el O(n²) completo.
        """
        from difflib import SequenceMatcher
        import collections

        # Índice invertido: token → [indices de filas en df]
        inv_idx: dict = collections.defaultdict(set)
        norms = df["titulo_norm"].tolist()
        token_sets = [set(n.split()) for n in norms]
        for idx, tokens in enumerate(token_sets):
            for tok in tokens:
                if len(tok) > 2:
                    inv_idx[tok].add(idx)

        # Solo comparar pares que comparten al menos 1 token
        visitados: set = set()
        grupos_dup: dict = {}  # titulo_norm → {indices}

        for i, t_i in enumerate(norms):
            candidatos = set()
            for tok in token_sets[i]:
                if len(tok) > 2:
                    candidatos |= inv_idx[tok]
            candidatos.discard(i)

            for j in candidatos:
                par = (min(i, j), max(i, j))
                if par in visitados:
                    continue
                visitados.add(par)

                t_j = norms[j]
                # quick_ratio como pre-filtro
                sm = SequenceMatcher(None, t_i, t_j, autojunk=False)
                if sm.quick_ratio() < umbral:
                    continue
                score = _similitud_jaccard(t_i, t_j)
                if score >= umbral:
                    clave = t_i  # usar el primer título normalizado como clave
                    grupos_dup.setdefault(clave, set()).add(i)
                    grupos_dup[clave].add(j)

        if not grupos_dup:
            return []

        # Construir DataFrame resultado
        filas = []
        for t_norm, indices in grupos_dup.items():
            sub = df.iloc[list(indices)]
            n_grupos = sub["grupo"].nunique()
            if n_grupos < 2:
                continue
            filas.append({
                "titulo_norm": t_norm,
                "titulo": sub["titulo"].iloc[0],
                "grupo": sub["grupo"].tolist(),
                "cedula": sub["cedula"].tolist(),
                "año": sub["año"].tolist(),
                "count": len(sub),
                "similitud_max": 1.0,  # al menos umbral (ya filtrado arriba)
                "match_por": "Título",
            })

        if not filas:
            return []

        agrup = pd.DataFrame(filas)
        return [agrup]


# =============================================================================
# ANÁLISIS DE CATEGORÍA 957 (datos GrupLAC)
# =============================================================================

class CategoriaAnalyzer957:
    """
    Analiza la posición de un grupo según los indicadores de Conv. 957,
    usando datame/gruplac_957.db como fuente de referencia histórica.

    Parámetros
    ----------
    db_path:
        Ruta a gruplac_957.db.
    percentil:
        Percentil (0-100) para calcular umbrales por categoría.
    area_conocimiento:
        Si se especifica, filtra los grupos de referencia por esta área
        al calcular umbrales estadísticos.
    """

    def __init__(
        self,
        db_path: str = "data/db/gruplac_957.db",
        percentil: int = 25,
        area_conocimiento: Optional[str] = None,
    ):
        self.db_path = db_path
        self.conn = sqlite3.connect(db_path)
        self.percentil = percentil
        self.area_conocimiento = area_conocimiento
        self.umbrales = self._calcular_umbrales()

    def cerrar(self):
        self.conn.close()

    def _calcular_umbrales(self) -> Dict:
        """
        Calcula el percentil de cada indicador por categoría.
        Usa numpy para los quantiles sobre arrays ya en memoria
        (la tabla grupos tiene solo ~56 filas, carga completa es OK).
        """
        import numpy as np

        # Verificar si existe columna de área
        area_filter = ""
        if self.area_conocimiento:
            cols = [
                r[1]
                for r in self.conn.execute("PRAGMA table_info(grupos)").fetchall()
            ]
            if "area_conocimiento" in cols:
                area_filter = f" AND area_conocimiento = '{self.area_conocimiento}'"

        # Carga mínima: solo las columnas necesarias
        indicadores_cols = [
            "ind_top", "ind_tipo_a", "ind_tipo_b",
            "ind_ap", "ind_dpc", "ind_fr_a", "ind_fr_b",
        ]
        sql = (
            f"SELECT categoria_asignada, {', '.join(indicadores_cols)} "
            f"FROM grupos WHERE categoria_asignada IN ('A','B','C','D'){area_filter}"
        )
        rows = self.conn.execute(sql).fetchall()
        if not rows:
            return {}

        # Agrupar por categoría con numpy para los cuantiles
        p = self.percentil / 100.0
        ind_keys = ["TOP", "TIPO_A", "TIPO_B", "AP", "DPC", "FR_A", "FR_B"]
        cat_data: Dict[str, list] = {c: [] for c in ["A", "B", "C", "D"]}
        for row in rows:
            cat = row[0]
            if cat in cat_data:
                cat_data[cat].append(row[1:])

        umbrales: Dict = {}
        for cat, data in cat_data.items():
            if not data:
                continue
            arr = np.array(data, dtype=float)
            umbrales[cat] = {
                k: float(np.nanpercentile(arr[:, i], self.percentil))
                for i, k in enumerate(ind_keys)
            }
        return umbrales

    def obtener_areas_disponibles(self) -> List[str]:
        """Devuelve las áreas de conocimiento disponibles en la BD."""
        cur = self.conn.cursor()
        cols = [
            row[1]
            for row in cur.execute("PRAGMA table_info(grupos)").fetchall()
        ]
        if "area_conocimiento" not in cols:
            return []
        rows = cur.execute(
            "SELECT DISTINCT area_conocimiento FROM grupos "
            "WHERE area_conocimiento IS NOT NULL ORDER BY area_conocimiento"
        ).fetchall()
        return [r[0] for r in rows]

    def listar_todos_grupos(self) -> List[Dict]:
        """Devuelve todos los grupos de gruplac_957.db (resultado cacheado)."""
        if not hasattr(self, "_cache_grupos"):
            cur = self.conn.cursor()
            cur.execute("SELECT * FROM grupos ORDER BY nombre_grupo")
            cols = [desc[0] for desc in cur.description]
            self._cache_grupos = [dict(zip(cols, row)) for row in cur.fetchall()]
            # Pre-calcular nombre normalizado para cada grupo
            for g in self._cache_grupos:
                g["_nombre_norm"] = _normalizar(g.get("nombre_grupo", ""))
                g["_tokens"] = set(g["_nombre_norm"].split())
        return self._cache_grupos

    def obtener_grupo_por_nombre(
        self, nombre_grupo: str, umbral_fuzzy: float = 0.60
    ) -> Optional[Dict]:
        """
        Matching progresivo con cache de grupos:
        1. Exacto  2. Normalizado exacto  3. quick_ratio pre-filtro + score difuso
        """
        from difflib import SequenceMatcher

        nombre_norm = _normalizar(nombre_grupo)
        tokens_q = set(nombre_norm.split())
        todos = self.listar_todos_grupos()

        # 1. Exacto
        for g in todos:
            if g.get("nombre_grupo") == nombre_grupo:
                return g

        # 2. Normalizado exacto
        for g in todos:
            if g["_nombre_norm"] == nombre_norm:
                return g

        # 3. Difuso con pre-filtro quick_ratio (evita SequenceMatcher completo en candidatos pobres)
        mejor_score = 0.0
        mejor_g = None
        sm = SequenceMatcher(autojunk=False)
        sm.set_seq2(nombre_norm)

        for g in todos:
            nb = g["_nombre_norm"]
            tokens_b = g["_tokens"]

            # Pre-filtro rápido: quick_ratio es una cota superior del ratio completo
            sm.set_seq1(nb)
            if sm.quick_ratio() < umbral_fuzzy * 0.8:
                continue

            seq = sm.ratio()
            if tokens_q and tokens_b:
                jac = len(tokens_q & tokens_b) / len(tokens_q | tokens_b)
            else:
                jac = 0.0
            score = 0.65 * seq + 0.35 * jac

            if score > mejor_score:
                mejor_score = score
                mejor_g = g

        return mejor_g if mejor_score >= umbral_fuzzy else None

    def analizar_grupo(self, nombre_grupo: str) -> Optional[Dict]:
        """
        Devuelve un análisis completo del grupo con indicadores actuales,
        umbrales de la siguiente categoría y brechas.
        """
        info = self.obtener_grupo_por_nombre(nombre_grupo)
        if not info:
            return None

        categoria_actual = info.get("categoria_asignada")
        ind_actuales = {
            "TOP":    info.get("ind_top", 0) or 0,
            "TIPO_A": info.get("ind_tipo_a", 0) or 0,
            "TIPO_B": info.get("ind_tipo_b", 0) or 0,
            "AP":     info.get("ind_ap", 0) or 0,
            "DPC":    info.get("ind_dpc", 0) or 0,
            "FR_A":   info.get("ind_fr_a", 0) or 0,
            "FR_B":   info.get("ind_fr_b", 0) or 0,
            "COHESION":     info.get("ind_cohesion", 0) or 0,
            "COLABORACION": info.get("ind_colaboracion", 0) or 0,
        }

        # Posición relativa frente al máximo histórico
        pos_relativa = {}
        for k, v in ind_actuales.items():
            maximo = info.get(f"max_{k.lower()}", 1) or 1
            pos_relativa[k] = v / maximo if maximo > 0 else 0.0

        # Categoría objetivo (próxima en la secuencia D→C→B→A)
        idx = (
            ORDEN_CATEGORIAS_MINCIENCIAS.index(categoria_actual)
            if categoria_actual in ORDEN_CATEGORIAS_MINCIENCIAS
            else -1
        )
        cat_objetivo = (
            ORDEN_CATEGORIAS_MINCIENCIAS[idx + 1]
            if 0 <= idx < len(ORDEN_CATEGORIAS_MINCIENCIAS) - 1
            else None
        )

        # Brechas
        brechas = {}
        if cat_objetivo and cat_objetivo in self.umbrales:
            for ind, umbral in self.umbrales[cat_objetivo].items():
                actual = ind_actuales.get(ind, 0)
                if actual < umbral:
                    brechas[ind] = {
                        "actual":     round(actual, 3),
                        "umbral":     round(umbral, 3),
                        "diferencia": round(umbral - actual, 3),
                    }

        cumple_para = {}
        if cat_objetivo and cat_objetivo in self.umbrales:
            req = self.umbrales[cat_objetivo]
            cumple_para[cat_objetivo] = all(
                ind_actuales.get(ind, 0) >= req[ind] for ind in req
            )

        return {
            "grupo_id":          info.get("id"),
            "categoria_actual":  categoria_actual,
            "categoria_objetivo": cat_objetivo,
            "area_conocimiento": info.get("area_conocimiento"),
            "indicadores":       ind_actuales,
            "posicion_relativa": pos_relativa,
            "umbrales":          self.umbrales.get(cat_objetivo, {}),
            "brechas":           brechas,
            "cumple_para":       cumple_para,
        }


# =============================================================================
# SIMULADOR DE CATEGORÍA CON DATOS INTERNOS
# =============================================================================

class SimuladorCategoriaInterna:
    """
    Simula la categoría 957 de un grupo usando la BD interna de la UTP.

    El simulador:
    1. Obtiene los integrantes del grupo desde la BD interna.
    2. Consulta sus productos dentro de las ventanas de MinCiencias.
    3. Mapea cada producto a un indicador 957 con su peso (lambda).
    4. Compara los indicadores resultantes con los umbrales estadísticos
       de la BD GrupLAC, opcionalmente filtrados por área de conocimiento.
    5. Calcula brechas y recomienda productos para subir de categoría.

    Parámetros
    ----------
    db_conn:
        Conexión SQLite a la BD interna (academia_utp_integrado.db).
    gruplac_db_path:
        Ruta a datame/gruplac_957.db para obtener umbrales de referencia.
    año_base:
        Año de cierre de la convocatoria (por defecto: año en curso).
    percentil:
        Percentil estadístico para los umbrales (por defecto: 25).
    """

    def __init__(
        self,
        db_conn,
        gruplac_db_path: str = "data/db/gruplac_957.db",
        año_base: Optional[int] = None,
        percentil: int = 25,
    ):
        self.conn = db_conn
        self.gruplac_db_path = gruplac_db_path
        self.año_base = año_base or datetime.now().year
        self.percentil = percentil
        self._umbrales_cache: Dict = {}

    def _obtener_umbrales(self, area: Optional[str] = None) -> Dict:
        """Carga o devuelve desde caché los umbrales estadísticos de GrupLAC."""
        clave = area or "__todas__"
        if clave in self._umbrales_cache:
            return self._umbrales_cache[clave]
        try:
            analizador = CategoriaAnalyzer957(
                db_path=self.gruplac_db_path,
                percentil=self.percentil,
                area_conocimiento=area,
            )
            self._umbrales_cache[clave] = analizador.umbrales
            analizador.cerrar()
        except Exception:
            self._umbrales_cache[clave] = {}
        return self._umbrales_cache[clave]

    def _cedulas_del_grupo(self, nombre_grupo: str) -> List[str]:
        """Devuelve las cédulas de los integrantes activos del grupo."""
        cur = self.conn.cursor()
        rows = cur.execute(
            "SELECT DISTINCT cedula FROM grupos WHERE grupo = ?",
            (nombre_grupo,),
        ).fetchall()
        return [r[0] for r in rows]

    def _calcular_indicadores(self, cedulas: List[str]) -> Dict[str, float]:
        """
        Calcula los valores de cada indicador 957 para las cédulas dadas,
        filtrando por las ventanas de evaluación de VENTANAS_957.
        """
        indicadores: Dict[str, float] = {k: 0.0 for k in INDICADORES_957}
        if not cedulas:
            return indicadores

        placeholders = ",".join("?" * len(cedulas))

        # ── Publicaciones (artículos, capítulos, libros) ──────────────────────
        filas_pub = self.conn.execute(
            f"""
            SELECT tipo, categoria, año
            FROM publicaciones
            WHERE cedula IN ({placeholders})
              AND año IS NOT NULL
            """,
            cedulas,
        ).fetchall()

        for tipo_prod, categoria, año in filas_pub:
            if not año:
                continue
            tipo_up = str(tipo_prod).upper() if tipo_prod else ""
            cat_up = str(categoria).upper().strip() if categoria else ""
            clave = None

            if tipo_up == "LIBRO":
                clave = "LIBRO"
            elif tipo_up == "CAPITULO":
                clave = "CAPITULO"
            elif cat_up in VENTANAS_957:
                clave = cat_up

            if clave and clave in VENTANAS_957:
                cfg = VENTANAS_957[clave]
                if self.año_base - int(año) <= cfg["ventana"]:
                    indicadores[cfg["indicador"]] += cfg["lambda"]

        # ── Trabajos de grado ─────────────────────────────────────────────────
        filas_tg = self.conn.execute(
            f"""
            SELECT programa, año
            FROM trabajos_grado
            WHERE cedula_director IN ({placeholders})
              AND año IS NOT NULL
              AND (calificacion IS NULL OR calificacion != 'NO CONDUCENTE')
            """,
            cedulas,
        ).fetchall()

        for programa, año in filas_tg:
            if not año:
                continue
            prog_norm = unidecode(str(programa).lower()) if programa else ""
            if "doctorado" in prog_norm:
                clave = "DOCTORADO"
            elif "maestr" in prog_norm:
                clave = "MAESTRIA"
            elif "especializ" in prog_norm:
                clave = "ESPECIALIZACION"
            else:
                clave = "PREGRADO"

            cfg = VENTANAS_957[clave]
            if self.año_base - int(año) <= cfg["ventana"]:
                indicadores[cfg["indicador"]] += cfg["lambda"]

        # ── Productos de innovación ───────────────────────────────────────────
        filas_inn = self.conn.execute(
            f"""
            SELECT tipo_producto, año
            FROM productos_innovacion
            WHERE cedula IN ({placeholders})
              AND año IS NOT NULL
            """,
            cedulas,
        ).fetchall()

        for tipo_prod, año in filas_inn:
            if not año:
                continue
            tipo_up = unidecode(str(tipo_prod).upper()) if tipo_prod else ""
            clave = None
            if "PATENTE" in tipo_up:
                clave = "PATENTE"
            elif "SOFTWARE" in tipo_up:
                clave = "SOFTWARE"
            elif "PROTOTIPO" in tipo_up:
                clave = "PROTOTIPO"

            if clave and clave in VENTANAS_957:
                cfg = VENTANAS_957[clave]
                if self.año_base - int(año) <= cfg["ventana"]:
                    indicadores[cfg["indicador"]] += cfg["lambda"]

        return indicadores

    def _estimar_categoria(
        self, indicadores: Dict[str, float], umbrales: Dict
    ) -> str:
        """
        Estima la categoría del grupo comparando sus indicadores
        contra los umbrales estadísticos de cada categoría.
        Devuelve la categoría más alta que el grupo supera.
        """
        for cat in reversed(ORDEN_CATEGORIAS_MINCIENCIAS):  # A, B, C, D
            if cat not in umbrales:
                continue
            if all(
                indicadores.get(ind, 0) >= val
                for ind, val in umbrales[cat].items()
            ):
                return cat
        return "Sin categoría"

    def analizar_grupo(
        self, nombre_grupo: str, area_conocimiento: Optional[str] = None
    ) -> Optional[Dict]:
        """
        Analiza un grupo de la BD interna y estima su categoría 957.

        Devuelve un diccionario con:
        - nombre_grupo, cedulas, num_integrantes
        - indicadores (valores calculados)
        - categoria_estimada
        - categoria_objetivo (siguiente categoría)
        - umbrales (umbrales de la categoría objetivo)
        - brechas (indicadores que no alcanzan el umbral)
        - recomendaciones (qué productos agregar para cerrar brechas)
        """
        cedulas = self._cedulas_del_grupo(nombre_grupo)
        if not cedulas:
            return None

        indicadores = self._calcular_indicadores(cedulas)
        umbrales = self._obtener_umbrales(area_conocimiento)

        cat_estimada = self._estimar_categoria(indicadores, umbrales)

        idx = (
            ORDEN_CATEGORIAS_MINCIENCIAS.index(cat_estimada)
            if cat_estimada in ORDEN_CATEGORIAS_MINCIENCIAS
            else -1
        )
        cat_objetivo = (
            ORDEN_CATEGORIAS_MINCIENCIAS[idx + 1]
            if 0 <= idx < len(ORDEN_CATEGORIAS_MINCIENCIAS) - 1
            else None
        )

        brechas = {}
        if cat_objetivo and cat_objetivo in umbrales:
            for ind, umbral in umbrales[cat_objetivo].items():
                actual = indicadores.get(ind, 0)
                if actual < umbral:
                    brechas[ind] = {
                        "actual":     round(actual, 3),
                        "umbral":     round(umbral, 3),
                        "diferencia": round(umbral - actual, 3),
                    }

        recomendaciones = self._generar_recomendaciones(brechas)

        return {
            "nombre_grupo":       nombre_grupo,
            "cedulas":            cedulas,
            "num_integrantes":    len(cedulas),
            "año_base":           self.año_base,
            "area_conocimiento":  area_conocimiento,
            "indicadores":        indicadores,
            "categoria_estimada": cat_estimada,
            "categoria_objetivo": cat_objetivo,
            "umbrales":           umbrales.get(cat_objetivo, {}),
            "brechas":            brechas,
            "recomendaciones":    recomendaciones,
        }

    @staticmethod
    def _generar_recomendaciones(brechas: Dict) -> List[Dict]:
        """
        Traduce brechas en recomendaciones concretas de productos.

        Mapeo indicador → tipo de producto sugerido y lambda unitario.
        """
        SUGERENCIAS = {
            "TOP": [
                {"producto": "Artículo A1 (Q1/Q2 SJR)", "lambda": 1.00},
                {"producto": "Libro resultado de investigación", "lambda": 2.00},
                {"producto": "Patente concedida", "lambda": 1.00},
            ],
            "TIPO_A": [
                {"producto": "Artículo A2 (Q3 SJR)", "lambda": 0.75},
                {"producto": "Capítulo de libro", "lambda": 1.00},
            ],
            "TIPO_B": [
                {"producto": "Artículo B (Q4 SJR o Colciencias B)", "lambda": 0.50},
                {"producto": "Software registrado", "lambda": 0.50},
                {"producto": "Prototipo", "lambda": 0.50},
            ],
            "AP": [
                {"producto": "Artículo C o D", "lambda": 0.25},
            ],
            "DPC": [
                {"producto": "Tesis doctoral dirigida", "lambda": 0.50},
            ],
            "FR_A": [
                {"producto": "Tesis de maestría dirigida", "lambda": 0.25},
            ],
            "FR_B": [
                {"producto": "Trabajo de grado pregrado/especialización", "lambda": 0.10},
            ],
        }

        recomendaciones = []
        for ind, info in brechas.items():
            diferencia = info["diferencia"]
            sugerencias_ind = SUGERENCIAS.get(ind, [])
            if not sugerencias_ind:
                continue
            mejor = sugerencias_ind[0]
            cantidad_minima = max(1, int(diferencia / mejor["lambda"]) + 1)
            recomendaciones.append({
                "indicador":   ind,
                "diferencia":  diferencia,
                "opciones": [
                    {
                        "producto":          s["producto"],
                        "lambda_unitario":   s["lambda"],
                        "cantidad_minima":   max(1, int(diferencia / s["lambda"]) + 1),
                    }
                    for s in sugerencias_ind
                ],
                "cantidad_optima": cantidad_minima,
                "producto_sugerido": mejor["producto"],
            })

        return sorted(recomendaciones, key=lambda r: r["diferencia"], reverse=True)

    def _calcular_indicadores_con_detalle(self, cedulas: List[str]) -> Dict:
        """
        Igual que _calcular_indicadores pero también retorna el listado de
        cada producto contabilizado en cada indicador.

        Retorna:
            {
              "indicadores": {ind: float},
              "productos_por_indicador": {ind: [{"titulo","año","tipo","lambda","fuente"}]}
            }
        """
        indicadores: Dict[str, float] = {k: 0.0 for k in INDICADORES_957}
        productos_por_ind: Dict[str, List[Dict]] = {k: [] for k in INDICADORES_957}

        if not cedulas:
            return {"indicadores": indicadores, "productos_por_indicador": productos_por_ind}

        placeholders = ",".join("?" * len(cedulas))

        # ── Publicaciones ──────────────────────────────────────────────────────
        filas_pub = self.conn.execute(
            f"""
            SELECT tipo, categoria, año, titulo
            FROM publicaciones
            WHERE cedula IN ({placeholders})
              AND año IS NOT NULL
            """,
            cedulas,
        ).fetchall()

        for tipo_prod, categoria, año, titulo in filas_pub:
            if not año:
                continue
            tipo_up = str(tipo_prod).upper() if tipo_prod else ""
            cat_up  = str(categoria).upper().strip() if categoria else ""
            clave   = None

            if tipo_up == "LIBRO":
                clave = "LIBRO"
            elif tipo_up == "CAPITULO":
                clave = "CAPITULO"
            elif cat_up in VENTANAS_957:
                clave = cat_up

            if clave and clave in VENTANAS_957:
                cfg = VENTANAS_957[clave]
                if self.año_base - int(año) <= cfg["ventana"]:
                    indicadores[cfg["indicador"]] += cfg["lambda"]
                    productos_por_ind[cfg["indicador"]].append({
                        "titulo":    titulo or "(sin título)",
                        "año":       año,
                        "tipo":      clave,
                        "categoria": categoria or "",
                        "lambda":    cfg["lambda"],
                        "fuente":    "publicaciones",
                    })

        # ── Trabajos de grado ──────────────────────────────────────────────────
        filas_tg = self.conn.execute(
            f"""
            SELECT programa, año, titulo
            FROM trabajos_grado
            WHERE cedula_director IN ({placeholders})
              AND año IS NOT NULL
              AND (calificacion IS NULL OR calificacion != 'NO CONDUCENTE')
            """,
            cedulas,
        ).fetchall()

        for programa, año, titulo in filas_tg:
            if not año:
                continue
            prog_norm = unidecode(str(programa).lower()) if programa else ""
            if "doctorado" in prog_norm:
                clave = "DOCTORADO"
            elif "maestr" in prog_norm:
                clave = "MAESTRIA"
            elif "especializ" in prog_norm:
                clave = "ESPECIALIZACION"
            else:
                clave = "PREGRADO"

            cfg = VENTANAS_957[clave]
            if self.año_base - int(año) <= cfg["ventana"]:
                indicadores[cfg["indicador"]] += cfg["lambda"]
                productos_por_ind[cfg["indicador"]].append({
                    "titulo":    titulo or "(sin título)",
                    "año":       año,
                    "tipo":      clave,
                    "categoria": programa or "",
                    "lambda":    cfg["lambda"],
                    "fuente":    "trabajos_grado",
                })

        # ── Productos de innovación ───────────────────────────────────────────
        filas_inn = self.conn.execute(
            f"""
            SELECT tipo_producto, año, nombre
            FROM productos_innovacion
            WHERE cedula IN ({placeholders})
              AND año IS NOT NULL
            """,
            cedulas,
        ).fetchall()

        for tipo_prod, año, nombre in filas_inn:
            if not año:
                continue
            tipo_up = unidecode(str(tipo_prod).upper()) if tipo_prod else ""
            clave = None
            if "PATENTE" in tipo_up:
                clave = "PATENTE"
            elif "SOFTWARE" in tipo_up:
                clave = "SOFTWARE"
            elif "PROTOTIPO" in tipo_up:
                clave = "PROTOTIPO"

            if clave and clave in VENTANAS_957:
                cfg = VENTANAS_957[clave]
                if self.año_base - int(año) <= cfg["ventana"]:
                    indicadores[cfg["indicador"]] += cfg["lambda"]
                    productos_por_ind[cfg["indicador"]].append({
                        "titulo":    nombre or "(sin nombre)",
                        "año":       año,
                        "tipo":      clave,
                        "categoria": tipo_prod or "",
                        "lambda":    cfg["lambda"],
                        "fuente":    "productos_innovacion",
                    })

        return {"indicadores": indicadores, "productos_por_indicador": productos_por_ind}

    def simular_escenario(
        self,
        nombre_grupo: str,
        productos_adicionales: Dict[str, int],
        area_conocimiento: Optional[str] = None,
    ) -> Optional[Dict]:
        """
        Simula qué categoría alcanzaría el grupo si agregara los productos
        especificados en productos_adicionales.

        productos_adicionales: {clave_VENTANAS_957: cantidad_a_agregar}
        Ejemplo: {"A1": 2, "MAESTRIA": 3}
        """
        base = self.analizar_grupo(nombre_grupo, area_conocimiento)
        if not base:
            return None

        ind_simulados = dict(base["indicadores"])
        for clave, cantidad in productos_adicionales.items():
            cfg = VENTANAS_957.get(clave.upper())
            if cfg:
                ind_simulados[cfg["indicador"]] += cfg["lambda"] * cantidad

        umbrales = self._obtener_umbrales(area_conocimiento)
        cat_simulada = self._estimar_categoria(ind_simulados, umbrales)

        return {
            "nombre_grupo":         nombre_grupo,
            "categoria_original":   base["categoria_estimada"],
            "categoria_simulada":   cat_simulada,
            "productos_adicionales": productos_adicionales,
            "indicadores_originales": base["indicadores"],
            "indicadores_simulados":  ind_simulados,
            "mejora": cat_simulada != base["categoria_estimada"],
        }

    # ── Plan de mejora 957: brechas por categoría y área de conocimiento ────

    def _cargar_referencia_957(self) -> List[Dict]:
        """
        Carga (con caché) la posición categoría/área de los grupos, combinando
        data/output/medicion_957.xlsx (75 grupos del documento oficial) y
        data/cache/categorias_grupos_957.json (125 grupos, respaldo).
        """
        if hasattr(self, "_ref_957"):
            return self._ref_957

        base = obtener_directorio_base()
        grupos_ref: List[Dict] = []
        vistos = set()

        try:
            ruta_medicion = base / "data" / "output" / "medicion_957.xlsx"
            df = pd.read_excel(ruta_medicion, sheet_name="grupos")
            for _, row in df.iterrows():
                nombre = row.get("nombre_grupo")
                if pd.isna(nombre):
                    continue
                categoria = row.get("categoria_confirmada")
                if pd.isna(categoria) or not categoria:
                    categoria = row.get("categoria")
                area = row.get("area_conocimiento")
                if pd.isna(categoria) or not categoria:
                    continue
                nombre = str(nombre)
                clave = _normalizar(nombre)
                grupos_ref.append({
                    "nombre": nombre,
                    "categoria": str(categoria).strip(),
                    "area": None if pd.isna(area) else str(area).strip(),
                })
                vistos.add(clave)
        except Exception:
            pass

        try:
            ruta_cache = base / "data" / "cache" / "categorias_grupos_957.json"
            with open(ruta_cache, "r", encoding="utf-8") as f:
                data = json.load(f)
            for nombre, info in data.get("grupos", {}).items():
                categoria = info.get("categoria_asignada")
                if not categoria:
                    continue
                clave = _normalizar(nombre)
                if clave in vistos:
                    continue
                grupos_ref.append({
                    "nombre": nombre,
                    "categoria": str(categoria).strip(),
                    "area": info.get("area_conocimiento"),
                })
                vistos.add(clave)
        except Exception:
            pass

        for g in grupos_ref:
            g["_norm"] = _normalizar(g["nombre"])
            g["_tokens"] = set(g["_norm"].split())

        self._ref_957 = grupos_ref
        return grupos_ref

    @staticmethod
    def _mejor_coincidencia(
        nombre: str, candidatos: List[Dict], umbral_fuzzy: float = 0.60, min_jaccard: float = 0.0
    ) -> Optional[Dict]:
        """
        Busca en `candidatos` (cada uno con claves '_norm' y '_tokens') el que
        mejor coincide con `nombre`: exacto normalizado, o difuso
        (SequenceMatcher + Jaccard de tokens).

        `min_jaccard`: para coincidencias difusas (no exactas), exige que la
        proporción de tokens compartidos sea al menos este valor. Evita falsos
        positivos entre nombres largos tipo "GRUPO DE INVESTIGACIÓN ..." que
        comparten muchas palabras genéricas pero corresponden a grupos
        distintos (p.ej. por mojibake del nombre, la coincidencia exacta
        falla y el difuso encuentra otro grupo cualquiera).
        """
        from difflib import SequenceMatcher

        nombre_norm = _normalizar(nombre)
        tokens_q = set(nombre_norm.split())

        for c in candidatos:
            if c["_norm"] == nombre_norm:
                return c

        mejor_score, mejor_c = 0.0, None
        sm = SequenceMatcher(autojunk=False)
        sm.set_seq2(nombre_norm)
        for c in candidatos:
            sm.set_seq1(c["_norm"])
            if sm.quick_ratio() < umbral_fuzzy * 0.8:
                continue
            tokens_c = c["_tokens"]
            jac = (
                len(tokens_q & tokens_c) / len(tokens_q | tokens_c)
                if tokens_q and tokens_c else 0.0
            )
            if jac < min_jaccard:
                continue
            seq = sm.ratio()
            score = 0.65 * seq + 0.35 * jac
            if score > mejor_score:
                mejor_score, mejor_c = score, c

        return mejor_c if mejor_score >= umbral_fuzzy else None

    def _buscar_grupo_interno(self, nombre_grupo: str) -> Optional[str]:
        """Devuelve el nombre exacto (tal como está en la BD interna) del
        grupo que mejor coincide con `nombre_grupo`, o None si no hay match."""
        if not hasattr(self, "_grupos_internos"):
            cur = self.conn.cursor()
            rows = cur.execute(
                "SELECT DISTINCT grupo FROM grupos WHERE grupo IS NOT NULL AND grupo != ''"
            ).fetchall()
            self._grupos_internos = [
                {"nombre": r[0], "_norm": _normalizar(r[0]), "_tokens": set(_normalizar(r[0]).split())}
                for r in rows
            ]

        encontrado = self._mejor_coincidencia(nombre_grupo, self._grupos_internos)
        return encontrado["nombre"] if encontrado else None

    # ── Datos oficiales (medicion_957.xlsx → hojas "indicadores"/"productos") ──

    def _cargar_datos_oficiales_957(self) -> Dict:
        """
        Carga (con caché) las hojas "indicadores", "productos" y "cuartiles"
        de medicion_957.xlsx: los valores oficiales de MinCiencias, ya
        calculados y categorizados a partir de los productos del documento
        de medición ("carpeta pdf").

        Retorna un dict con:
            "indicadores": {nombre_normalizado: {indicador_957: valor_indicador}}
            "productos":   {nombre_normalizado: {seccion: {subtipo: total}}}
            "cuartiles":   {nombre_normalizado: {indicador_957: {min,q4,q3,q2,max,valor_grupo}}}
            "ventanas_subtipo": {subtipo: {"ventana": int, "seccion": str}} — la
                ventana de vigencia (años) de cada subtipo de producto, tal
                como aparece en la hoja "productos" (es la misma para todos
                los grupos: ART_*=7, LIB_*/CAP_LIB_*=5, etc.).
            "candidatos":  lista para _mejor_coincidencia (nombre/_norm/_tokens)
        """
        if hasattr(self, "_oficial_957"):
            return self._oficial_957

        indicadores_por_grupo: Dict[str, Dict[str, float]] = {}
        productos_por_grupo: Dict[str, Dict[str, Dict[str, int]]] = {}
        cuartiles_por_grupo: Dict[str, Dict[str, Dict[str, float]]] = {}
        ventanas_subtipo: Dict[str, Dict[str, object]] = {}
        nombres: Dict[str, str] = {}

        CUARTIL_A_INDICADOR = {label: ind for ind, label in INDICADOR_957_A_CUARTIL.items()}

        try:
            ruta = obtener_directorio_base() / "data" / "output" / "medicion_957.xlsx"

            df_i = pd.read_excel(ruta, sheet_name="indicadores")
            for _, row in df_i.iterrows():
                grupo = row.get("grupo")
                codigo = row.get("indicador")
                valor = row.get("valor_indicador")
                if pd.isna(grupo) or codigo not in CODIGO_957_A_INDICADOR or pd.isna(valor):
                    continue
                clave = _normalizar(str(grupo))
                nombres.setdefault(clave, str(grupo))
                indicadores_por_grupo.setdefault(clave, {})[CODIGO_957_A_INDICADOR[codigo]] = float(valor)

            df_p = pd.read_excel(ruta, sheet_name="productos")
            for _, row in df_p.iterrows():
                grupo = row.get("grupo")
                seccion = row.get("seccion")
                subtipo = row.get("subtipo")
                total = row.get("total")
                if pd.isna(grupo) or pd.isna(seccion) or pd.isna(subtipo) or pd.isna(total):
                    continue
                clave = _normalizar(str(grupo))
                nombres.setdefault(clave, str(grupo))
                sub_dict = productos_por_grupo.setdefault(clave, {}).setdefault(str(seccion), {})
                sub_dict[str(subtipo)] = sub_dict.get(str(subtipo), 0) + int(total)

                subtipo_s = str(subtipo)
                if subtipo_s not in ventanas_subtipo:
                    ventana_val = row.get("ventana")
                    ventanas_subtipo[subtipo_s] = {
                        "ventana": int(ventana_val) if pd.notna(ventana_val) and ventana_val else 5,
                        "seccion": str(seccion),
                    }

            df_c = pd.read_excel(ruta, sheet_name="cuartiles")
            for _, row in df_c.iterrows():
                grupo = row.get("grupo")
                cuartil_label = row.get("cuartil")
                if pd.isna(grupo) or cuartil_label not in CUARTIL_A_INDICADOR:
                    continue
                clave = _normalizar(str(grupo))
                nombres.setdefault(clave, str(grupo))
                ind = CUARTIL_A_INDICADOR[cuartil_label]
                cuartiles_por_grupo.setdefault(clave, {})[ind] = {
                    "min": float(row.get("min")) if pd.notna(row.get("min")) else None,
                    "q4":  float(row.get("q4"))  if pd.notna(row.get("q4"))  else None,
                    "q3":  float(row.get("q3"))  if pd.notna(row.get("q3"))  else None,
                    "q2":  float(row.get("q2"))  if pd.notna(row.get("q2"))  else None,
                    "max": float(row.get("max")) if pd.notna(row.get("max")) else None,
                    "valor_grupo": float(row.get("valor_grupo")) if pd.notna(row.get("valor_grupo")) else None,
                }
        except Exception:
            pass

        candidatos = [
            {"nombre": nombre, "_norm": clave, "_tokens": set(clave.split())}
            for clave, nombre in nombres.items()
        ]

        self._oficial_957 = {
            "indicadores": indicadores_por_grupo,
            "productos": productos_por_grupo,
            "cuartiles": cuartiles_por_grupo,
            "ventanas_subtipo": ventanas_subtipo,
            "candidatos": candidatos,
        }
        return self._oficial_957

    def _indicadores_oficiales(self, nombre_grupo: str) -> Optional[Dict[str, float]]:
        """
        Valores oficiales (hoja "indicadores" de medicion_957.xlsx) del grupo
        que mejor coincide con `nombre_grupo`, mapeados a las claves de
        INDICADORES_957. None si el grupo no está entre los cubiertos por el
        documento de medición.
        """
        datos = self._cargar_datos_oficiales_957()
        encontrado = self._mejor_coincidencia(nombre_grupo, datos["candidatos"], min_jaccard=0.5)
        if not encontrado:
            return None
        return datos["indicadores"].get(encontrado["_norm"])

    def _productos_oficiales(self, nombre_grupo: str) -> Dict[str, Dict[str, int]]:
        """
        Productos oficiales (hoja "productos" de medicion_957.xlsx) del grupo
        que mejor coincide con `nombre_grupo`, agrupados por sección y
        subtipo: {seccion: {subtipo: total}}. {} si no hay datos.
        """
        datos = self._cargar_datos_oficiales_957()
        encontrado = self._mejor_coincidencia(nombre_grupo, datos["candidatos"], min_jaccard=0.5)
        if not encontrado:
            return {}
        return datos["productos"].get(encontrado["_norm"], {})

    def _cuartiles_oficiales(self, nombre_grupo: str) -> Optional[Dict[str, Dict[str, float]]]:
        """
        Distribución nacional (hoja "cuartiles" de medicion_957.xlsx) de cada
        indicador 957 para la categoría/área del grupo que mejor coincide con
        `nombre_grupo`: {indicador_957: {min,q4,q3,q2,max,valor_grupo}}.

        Estos valores son los mismos para todos los grupos que comparten
        categoría y área de conocimiento (representan la distribución
        nacional de ese indicador para esa combinación), por lo que sirven
        como "ecuación" de MinCiencias: el valor mínimo que un indicador debe
        alcanzar para contar hacia una categoría se obtiene cruzando esta
        distribución con CUARTIL_OBJETIVO_POR_CATEGORIA.

        None si el grupo no está entre los cubiertos por el documento de
        medición.
        """
        datos = self._cargar_datos_oficiales_957()
        encontrado = self._mejor_coincidencia(nombre_grupo, datos["candidatos"], min_jaccard=0.5)
        if not encontrado:
            return None
        return datos["cuartiles"].get(encontrado["_norm"])

    @staticmethod
    def _etiqueta_subtipo(subtipo: str, seccion: str = "") -> str:
        """Traduce el código de subtipo de producto (hoja "productos" de
        medicion_957.xlsx) a una etiqueta legible. Para códigos sin una
        traducción conocida se usa una etiqueta genérica "<sección> — <código>"
        en vez de inventar un nombre oficial."""
        ETIQUETAS = {
            "ART_A1": "Artículo Q1 (A1)", "ART_OPEN_A1": "Artículo Q1 Open Access (A1)",
            "ART_A2": "Artículo Q2 (A2)", "ART_OPEN_A2": "Artículo Q2 Open Access (A2)",
            "ART_B":  "Artículo Q3 (B)", "ART_OPEN_B":  "Artículo Q3 Open Access (B)",
            "ART_C":  "Artículo Q4 (C)", "ART_OPEN_C":  "Artículo Q4 Open Access (C)",
            "ART_D":  "Artículo D",      "ART_OPEN_D":  "Artículo D Open Access",
            "LIB_A1": "Libro A1", "LIB_A": "Libro A", "LIB_B": "Libro B", "LIB_C": "Libro C",
            "CAP_LIB_A1": "Capítulo de libro A1", "CAP_LIB_A": "Capítulo de libro A",
            "CAP_LIB_B": "Capítulo de libro B", "CAP_LIB_C": "Capítulo de libro C",
            "PATENTE": "Patente", "SOFTWARE": "Software registrado", "PROTOTIPO": "Prototipo",
            "DOCTORADO": "Tesis de doctorado dirigida",
            "MAESTRIA": "Tesis de maestría dirigida",
            "ESPECIALIZACION": "Trabajo de especialización dirigido",
            "PREGRADO": "Trabajo de grado de pregrado dirigido",
        }
        if subtipo in ETIQUETAS:
            return ETIQUETAS[subtipo]
        SECCION_LABEL = {
            "NC_TOP": "Nuevo conocimiento TOP", "NC_A": "Nuevo conocimiento A",
            "NC_B": "Nuevo conocimiento B", "ASC": "Apropiación social",
            "DPC": "Divulgación pública",
        }
        prefijo = SECCION_LABEL.get(seccion)
        return f"{prefijo} — {subtipo}" if prefijo else subtipo.replace("_", " ").title()

    def opciones_simulador(
        self,
        nombre_grupo: str,
        requisitos_minimos_objetivo: Optional[Dict[str, Dict]],
        indicadores_actuales: Optional[Dict[str, float]],
    ) -> Dict[str, Dict]:
        """
        Para cada indicador con brecha frente a `requisitos_minimos_objetivo`
        (cumple == False), calcula las opciones del simulador "¿qué productos
        me faltan?" usando ÚNICAMENTE datos de medicion_957.xlsx:

        - "ventanas_subtipo" (hoja "productos"): ventana de vigencia (años) de
          cada subtipo, igual para todos los grupos.
        - λ de cada subtipo = ln(1 + total/ventana), la MISMA fórmula que la
          columna "lambda_val" de la hoja "productos" (verificada contra los
          ~1500 registros de esa hoja).
        - "ratio" = tu valor oficial del indicador (hoja "indicadores") ÷ tu λ
          total actual (suma de λ de todos los subtipos de esa sección que ya
          tienes). Este factor es propio de TU grupo —no un valor inventado—
          y permite proyectar, para cantidades hipotéticas de cada producto,
          un "valor simulado" en la MISMA escala que "Tu valor"/"Requisito
          mínimo" de la sección 4 (sin el desajuste de escala de versiones
          anteriores del simulador).

        Para FR_A/FR_B (formación de recurso humano) el documento oficial no
        desglosa productos: se devuelve {"sin_desglose": True}.

        {indicador: {
            "seccion": str,
            "ratio": float | None,
            "lambda_actual_total": float,
            "lambda_objetivo_total": float | None,
            "productos": [{"subtipo", "producto", "ventana", "total_actual",
                            "lambda_actual"}, ...]  (máx. 6, ordenados por
                            impacto marginal de la próxima unidad),
            "mas_disponibles": int,
        }}
        """
        if not requisitos_minimos_objetivo:
            return {}

        datos = self._cargar_datos_oficiales_957()
        ventanas_subtipo = datos.get("ventanas_subtipo", {})
        productos_propios = self._productos_oficiales(nombre_grupo)
        indicadores_actuales = indicadores_actuales or {}

        resultado: Dict[str, Dict] = {}
        for ind, fila in requisitos_minimos_objetivo.items():
            if fila.get("cumple") or fila.get("requisito_minimo") is None:
                continue

            seccion = SECCION_957_POR_INDICADOR.get(ind)
            if seccion is None:
                resultado[ind] = {"sin_desglose": True}
                continue

            propios = productos_propios.get(seccion, {})
            subtipos_seccion = [
                s for s, cfg in ventanas_subtipo.items() if cfg["seccion"] == seccion
            ]

            lambda_actual_total = 0.0
            candidatos = []
            for subtipo in subtipos_seccion:
                ventana = ventanas_subtipo[subtipo]["ventana"]
                total_actual = propios.get(subtipo, 0)
                lam_actual = math.log(1 + total_actual / ventana)
                lambda_actual_total += lam_actual
                marginal = math.log(1 + (total_actual + 1) / ventana) - lam_actual
                candidatos.append({
                    "subtipo": subtipo,
                    "producto": self._etiqueta_subtipo(subtipo, seccion),
                    "ventana": ventana,
                    "total_actual": total_actual,
                    "lambda_actual": lam_actual,
                    "_marginal": marginal,
                })

            valor_actual = indicadores_actuales.get(ind, 0.0)
            ratio = (valor_actual / lambda_actual_total) if lambda_actual_total > 0 else None
            requisito = fila["requisito_minimo"]
            lambda_objetivo_total = (requisito / ratio) if ratio else None

            # Primero los tipos de producto que el grupo YA genera (más
            # realista: "produce más de lo mismo"), ordenados por impacto
            # marginal de la próxima unidad; luego tipos nuevos, en el mismo
            # orden.
            candidatos.sort(key=lambda c: (c["total_actual"] > 0, c["_marginal"]), reverse=True)
            for c in candidatos:
                c.pop("_marginal")

            resultado[ind] = {
                "seccion": seccion,
                "ratio": ratio,
                "lambda_actual_total": lambda_actual_total,
                "lambda_objetivo_total": lambda_objetivo_total,
                "productos": candidatos[:6],
                "mas_disponibles": max(0, len(candidatos) - 6),
            }
        return resultado

    def _generar_recomendaciones_oficial(
        self, nombre_grupo: str, brechas: Dict, nombre_ref: Optional[str]
    ) -> List[Dict]:
        """
        Traduce brechas en recomendaciones concretas comparando, producto por
        producto (hoja "productos" de medicion_957.xlsx), lo que ya tiene el
        grupo analizado contra lo que tiene el grupo de referencia más débil
        que ya alcanzó la categoría objetivo.

        Para los indicadores sin desglose de productos en el documento
        oficial (FR_A/FR_B, formación de recurso humano) se usan las
        sugerencias genéricas de _generar_recomendaciones como aproximación.
        """
        productos_propios = self._productos_oficiales(nombre_grupo)
        productos_ref = self._productos_oficiales(nombre_ref) if nombre_ref else {}

        recomendaciones = []
        for ind, info in brechas.items():
            diferencia = info["diferencia"]
            seccion = SECCION_957_POR_INDICADOR.get(ind)
            opciones: List[Dict] = []
            if seccion:
                propios = productos_propios.get(seccion, {})
                referencia = productos_ref.get(seccion, {})
                for subtipo in sorted(set(propios) | set(referencia)):
                    tot_propio = propios.get(subtipo, 0)
                    tot_ref = referencia.get(subtipo, 0)
                    falta = tot_ref - tot_propio
                    if falta > 0:
                        opciones.append({
                            "producto": self._etiqueta_subtipo(subtipo),
                            "lambda_unitario": float(tot_propio),
                            "cantidad_minima": int(falta),
                        })
                opciones.sort(key=lambda o: o["cantidad_minima"], reverse=True)

            if not opciones:
                # Sin productos del grupo de referencia a igualar: TOP/TIPO_A/TIPO_B/AP/DPC
                # ya tienen tantos o más productos de cada subtipo, pero el índice oficial
                # sigue marcando brecha (probablemente por el factor de antigüedad λ de
                # cada producto). FR_A/FR_B no tienen desglose de productos en el
                # documento oficial.
                mensaje = (
                    "MinCiencias no desglosa este indicador por tipo de producto en el "
                    "documento oficial (formación de recurso humano)."
                    if not seccion else
                    "Ya tienes tantos o más productos de cada tipo que el grupo de "
                    "referencia; la brecha del índice puede deberse a la antigüedad (λ) "
                    "de tus productos — revisa la pestaña 'Datos paper 957'."
                )
                opciones = [{
                    "producto": mensaje,
                    "lambda_unitario": None,
                    "cantidad_minima": None,
                }]

            recomendaciones.append({
                "indicador": ind,
                "diferencia": diferencia,
                "opciones": opciones,
                "cantidad_optima": opciones[0]["cantidad_minima"],
                "producto_sugerido": opciones[0]["producto"],
            })

        return sorted(recomendaciones, key=lambda r: r["diferencia"], reverse=True)

    def analizar_brechas_area(
        self, nombre_grupo: str, grupo_referencia_manual: Optional[str] = None
    ) -> Optional[Dict]:
        """
        Identifica qué le falta a un grupo para subir a la siguiente categoría
        957 (D→C→B→A→A1).

        Fuente de los indicadores (tanto del grupo analizado como de los
        grupos de referencia):
        - "oficial": valores ya calculados por MinCiencias en la hoja
          "indicadores" de medicion_957.xlsx (a partir de los productos del
          documento de medición). Se usa siempre que el grupo esté entre los
          cubiertos por ese documento.
        - "bd_interna": si el grupo no aparece en medicion_957.xlsx, se
          recalculan los indicadores desde la BD interna
          (academia_utp_integrado.db) según VENTANAS_957.

        Los grupos de referencia (misma área de conocimiento, ya en la
        categoría objetivo) se evalúan con la MISMA fuente que el grupo
        analizado, para no mezclar escalas. El umbral de cada indicador es el
        percentil `self.percentil` de esos grupos de referencia. Las brechas
        resultantes se traducen en recomendaciones de productos mediante
        _generar_recomendaciones_oficial (fuente "oficial") o
        _generar_recomendaciones (fuente "bd_interna").

        `grupo_referencia_manual`: si se indica (debe ser uno de los nombres
        devueltos en "grupos_referencia_disponibles" de una llamada previa,
        es decir, un grupo de la MISMA área de conocimiento que ya está en
        categoria_objetivo), se usa ese grupo —en vez del "más débil"
        automático— como punto de comparación en
        "grupo_referencia_manual" y, si fuente == "oficial", también como
        referencia para las recomendaciones de productos.

        Cuando fuente == "oficial" también se calcula
        "requisitos_minimos_objetivo": para cada indicador, la distribución
        nacional (min/q4/q3/q2/max, hoja "cuartiles") de los grupos en
        categoria_objetivo del área, y el valor mínimo que MinCiencias exige
        para esa categoría (columna dada por CUARTIL_OBJETIVO_POR_CATEGORIA),
        comparado con el valor actual del grupo. Esta es la "ecuación" oficial
        de MinCiencias para la categoría objetivo.
        """
        referencias = self._cargar_referencia_957()
        ref = self._mejor_coincidencia(nombre_grupo, referencias)
        if not ref:
            return {
                "nombre_grupo": nombre_grupo,
                "categoria_actual": None,
                "area_conocimiento": None,
                "categoria_objetivo": None,
                "mensaje": "No se encontró información de categoría/área para este grupo "
                           "ni en medicion_957.xlsx ni en categorias_grupos_957.json.",
            }

        categoria_actual = ref["categoria"]
        area = ref["area"]

        if categoria_actual not in ORDEN_CATEGORIAS_MINCIENCIAS:
            return {
                "nombre_grupo":      nombre_grupo,
                "categoria_actual":  categoria_actual,
                "area_conocimiento": area,
                "categoria_objetivo": None,
                "mensaje": f"La categoría '{categoria_actual}' no hace parte de la escala D-C-B-A-A1.",
            }

        idx = ORDEN_CATEGORIAS_MINCIENCIAS.index(categoria_actual)
        if idx == len(ORDEN_CATEGORIAS_MINCIENCIAS) - 1:
            return {
                "nombre_grupo":      nombre_grupo,
                "categoria_actual":  categoria_actual,
                "area_conocimiento": area,
                "categoria_objetivo": None,
                "mensaje": "El grupo ya está en la categoría máxima (A1).",
            }
        categoria_objetivo = ORDEN_CATEGORIAS_MINCIENCIAS[idx + 1]

        # ── Indicadores del grupo analizado: oficial (medicion_957.xlsx) ────
        # si está cubierto por el documento de medición; si no, BD interna.
        indicadores_oficiales = self._indicadores_oficiales(nombre_grupo)
        if indicadores_oficiales:
            fuente = "oficial"
            indicadores_actuales = {k: indicadores_oficiales.get(k, 0.0) for k in INDICADORES_957}
        else:
            fuente = "bd_interna"
            cedulas = self._cedulas_del_grupo(nombre_grupo)
            indicadores_actuales = self._calcular_indicadores(cedulas)

        # Grupos de referencia: misma área de conocimiento, ya en categoria_objetivo
        area_norm = _normalizar(area) if area else None
        refs_objetivo = [
            g for g in referencias
            if g["categoria"] == categoria_objetivo
            and g.get("area")
            and (area_norm is None or _normalizar(g["area"]) == area_norm)
            and _normalizar(g["nombre"]) != _normalizar(nombre_grupo)
        ]

        manual_norm = _normalizar(grupo_referencia_manual) if grupo_referencia_manual else None

        valores_por_indicador: Dict[str, List[float]] = {k: [] for k in INDICADORES_957}
        grupos_referencia_usados: List[str] = []
        grupo_minimo: Optional[Dict] = None   # el grupo objetivo "más débil" del área (menor suma de indicadores)
        grupo_similar: Optional[Dict] = None  # el grupo objetivo de perfil más parecido al del grupo analizado
        grupo_manual: Optional[Dict] = None   # el grupo elegido manualmente por el usuario
        for g in refs_objetivo:
            if fuente == "oficial":
                ind_oficial_g = self._indicadores_oficiales(g["nombre"])
                if not ind_oficial_g:
                    continue
                ind_g = {k: ind_oficial_g.get(k, 0.0) for k in INDICADORES_957}
            else:
                nombre_interno = self._buscar_grupo_interno(g["nombre"])
                if not nombre_interno:
                    continue
                ceds_g = self._cedulas_del_grupo(nombre_interno)
                if not ceds_g:
                    continue
                ind_g = self._calcular_indicadores(ceds_g)

            for k, v in ind_g.items():
                valores_por_indicador[k].append(v)
            grupos_referencia_usados.append(g["nombre"])

            total_g = sum(ind_g.values())
            if grupo_minimo is None or total_g < grupo_minimo["total"]:
                grupo_minimo = {"nombre": g["nombre"], "indicadores": ind_g, "total": total_g}

            distancia = sum(
                (ind_g.get(k, 0) - indicadores_actuales.get(k, 0)) ** 2
                for k in INDICADORES_957
            )
            if grupo_similar is None or distancia < grupo_similar["distancia"]:
                grupo_similar = {"nombre": g["nombre"], "indicadores": ind_g, "distancia": distancia}

            if manual_norm is not None and _normalizar(g["nombre"]) == manual_norm:
                grupo_manual = {"nombre": g["nombre"], "indicadores": ind_g}

        umbrales_referencia: Dict[str, float] = {}
        for k, valores in valores_por_indicador.items():
            if valores:
                umbrales_referencia[k] = float(pd.Series(valores).quantile(self.percentil / 100.0))

        brechas = {}
        for ind, umbral in umbrales_referencia.items():
            actual = indicadores_actuales.get(ind, 0)
            if actual < umbral:
                brechas[ind] = {
                    "actual":     round(actual, 3),
                    "umbral":     round(umbral, 3),
                    "diferencia": round(umbral - actual, 3),
                }

        # El grupo manual, si se indicó y se encontró, reemplaza al "más débil"
        # como referencia para las recomendaciones de productos.
        nombre_ref_recomendacion = (
            grupo_manual["nombre"] if grupo_manual
            else grupo_minimo["nombre"] if grupo_minimo
            else None
        )
        if fuente == "oficial":
            recomendaciones = self._generar_recomendaciones_oficial(nombre_grupo, brechas, nombre_ref_recomendacion)
        else:
            recomendaciones = self._generar_recomendaciones(brechas)

        mensaje = None
        if not grupos_referencia_usados:
            if fuente == "oficial":
                mensaje = (
                    f"No se encontraron grupos de referencia con datos oficiales en "
                    f"medicion_957.xlsx para la categoría '{categoria_objetivo}' y el área '{area}'."
                )
            else:
                mensaje = (
                    f"No se encontraron grupos internos de referencia en categoría "
                    f"'{categoria_objetivo}' para el área '{area}'."
                )

        def _resumen_grupo_ref(g: Optional[Dict]) -> Optional[Dict]:
            if g is None:
                return None
            diferencias = {}
            for ind in INDICADORES_957:
                actual = indicadores_actuales.get(ind, 0)
                ref_val = g["indicadores"].get(ind, 0)
                diferencias[ind] = round(ref_val - actual, 3)
            return {
                "nombre":      g["nombre"],
                "indicadores": {k: round(v, 3) for k, v in g["indicadores"].items()},
                "diferencias": diferencias,
            }

        # ── Requisitos mínimos oficiales para categoria_objetivo (hoja "cuartiles") ──
        # Distribución nacional (min/q4/q3/q2/max) para grupos de la misma área que
        # ya están en categoria_objetivo, y el valor mínimo exigido por MinCiencias
        # para que cada indicador cuente hacia esa categoría.
        requisitos_minimos_objetivo: Optional[Dict[str, Dict]] = None
        grupo_referencia_cuartiles: Optional[str] = None
        if fuente == "oficial":
            for candidato in (grupo_manual, grupo_minimo, grupo_similar):
                if candidato:
                    cuartiles_obj = self._cuartiles_oficiales(candidato["nombre"])
                    if cuartiles_obj:
                        grupo_referencia_cuartiles = candidato["nombre"]
                        break
            else:
                cuartiles_obj = None

            columna_objetivo = CUARTIL_OBJETIVO_POR_CATEGORIA.get(categoria_objetivo)
            if cuartiles_obj and columna_objetivo:
                requisitos_minimos_objetivo = {}
                for ind in INDICADORES_957:
                    fila = cuartiles_obj.get(ind)
                    if not fila:
                        continue
                    requisito = fila.get(columna_objetivo)
                    actual = indicadores_actuales.get(ind, 0.0)
                    requisitos_minimos_objetivo[ind] = {
                        "min": fila.get("min"),
                        "q4":  fila.get("q4"),
                        "q3":  fila.get("q3"),
                        "q2":  fila.get("q2"),
                        "max": fila.get("max"),
                        "columna_objetivo": columna_objetivo,
                        "requisito_minimo": requisito,
                        "tu_valor": round(actual, 3),
                        "cumple": (actual >= requisito) if requisito is not None else None,
                    }

        return {
            "nombre_grupo":               nombre_grupo,
            "categoria_actual":           categoria_actual,
            "area_conocimiento":          area,
            "categoria_objetivo":         categoria_objetivo,
            "fuente_indicadores":         fuente,
            "indicadores_actuales":       indicadores_actuales,
            "umbrales_referencia":        {k: round(v, 3) for k, v in umbrales_referencia.items()},
            "num_grupos_referencia":      len(grupos_referencia_usados),
            "grupos_referencia":          grupos_referencia_usados,
            "grupos_referencia_disponibles": sorted(grupos_referencia_usados),
            "grupo_referencia_minimo":    _resumen_grupo_ref(grupo_minimo),
            "grupo_referencia_similar":   _resumen_grupo_ref(grupo_similar),
            "grupo_referencia_manual":    _resumen_grupo_ref(grupo_manual),
            "requisitos_minimos_objetivo": requisitos_minimos_objetivo,
            "grupo_referencia_cuartiles": grupo_referencia_cuartiles,
            "brechas":                    brechas,
            "recomendaciones":            recomendaciones,
            "mensaje":                    mensaje,
        }


# =============================================================================
# ANÁLISIS DE CONTRIBUCIÓN CRUZADA ENTRE GRUPOS
# =============================================================================

class ReporteContribucionCruzada:
    """
    Detecta productos coautorados por miembros de distintos grupos que
    podrían mejorar los indicadores de cohesión y colaboración (Conv. 957).

    Lógica MinCiencias 957:
    - ind_cohesion: productos donde ≥2 integrantes del MISMO grupo son coautores.
    - ind_colaboracion: productos con coautores externos al grupo.

    Si dos autores de grupos diferentes coautoron un paper y ese paper está
    registrado solo en el grupo A, el grupo B está perdiendo crédito de
    cohesión o colaboración.

    Parámetros
    ----------
    db_conn:
        Conexión SQLite a academia_utp_integrado.db.
    solo_activos:
        Si True, solo considera miembros cuyo tipo_miembro no indica retiro.
    """

    TIPOS_RETIRADO = {"ex integrante", "exintegrante", "retirado", "retirada", "inactivo"}

    def __init__(self, db_conn, solo_activos: bool = True):
        self.conn = db_conn
        self.solo_activos = solo_activos

    # ── utilidades ──────────────────────────────────────────────────────────

    def _integrantes_por_grupo(self) -> Dict[str, List[str]]:
        """Devuelve {nombre_grupo: [cedula, ...]} filtrando activos si aplica."""
        query = "SELECT cedula, grupo, tipo_miembro FROM grupos WHERE grupo IS NOT NULL"
        rows = self.conn.execute(query).fetchall()
        grupos: Dict[str, List[str]] = {}
        for cedula, grupo, tipo in rows:
            if self.solo_activos:
                tipo_norm = _normalizar(tipo or "")
                if any(t in tipo_norm for t in self.TIPOS_RETIRADO):
                    continue
            grupos.setdefault(grupo, []).append(cedula)
        return grupos

    def _cedula_a_grupos(self) -> Dict[str, List[str]]:
        """Inverso: {cedula: [grupos donde está activo]}."""
        por_grupo = self._integrantes_por_grupo()
        inverso: Dict[str, List[str]] = {}
        for grupo, cedulas in por_grupo.items():
            for c in cedulas:
                inverso.setdefault(c, []).append(grupo)
        return inverso

    @staticmethod
    def _tuplas_coautorias(conn) -> List[tuple]:
        """
        Retorna (titulo_norm, cedula, grupo, año, categoria) de publicaciones.
        Cada paper con N autores aparece N veces (una por cedula).
        """
        rows = conn.execute(
            """
            SELECT titulo, cedula, grupo, año, categoria, tipo
            FROM publicaciones
            WHERE titulo IS NOT NULL AND cedula IS NOT NULL
            """
        ).fetchall()
        result = []
        for titulo, cedula, grupo, año, cat, tipo in rows:
            tn = _normalizar(str(titulo))
            if len(tn) >= 10:
                result.append((tn, str(cedula), grupo or "", año, cat or "", tipo or ""))
        return result

    # ── análisis principal ───────────────────────────────────────────────────

    def analizar(self) -> Dict:
        """
        Detecta oportunidades de cohesión y colaboración.

        Retorna dict con:
          autores_multigrupo, oportunidades_cohesion,
          oportunidades_colaboracion, resumen_por_grupo
        """
        cedula_a_grupos = self._cedula_a_grupos()
        multigrupo = {c: gs for c, gs in cedula_a_grupos.items() if len(gs) > 1}

        coautorias = self._tuplas_coautorias(self.conn)
        if not coautorias:
            return {
                "autores_multigrupo": multigrupo,
                "oportunidades_cohesion": [],
                "oportunidades_colaboracion": [],
                "resumen_por_grupo": {},
            }

        # Agrupar por título normalizado
        papers: Dict[str, List] = {}
        for tn, cedula, grupo, año, cat, tipo in coautorias:
            papers.setdefault(tn, []).append(
                {"cedula": cedula, "grupo": grupo or "", "año": año, "cat": cat or ""}
            )

        # Solo papers con ≥2 autores distintos
        papers_multi = {
            t: auths for t, auths in papers.items()
            if len({a["cedula"] for a in auths}) >= 2
        }

        oport_cohesion: List[Dict] = []
        oport_colaboracion: List[Dict] = []
        resumen: Dict[str, Dict[str, int]] = {}

        for titulo_norm, autores in papers_multi.items():
            cedulas_paper = [a["cedula"] for a in autores]
            # Grupos donde está registrado el paper (excluir vacíos)
            grupos_paper = {a["grupo"] for a in autores if a["grupo"]}

            # Para cada cédula del paper, obtener todos sus grupos activos
            # Usar solo cedula_a_grupos (BD interna), NO el grupo del paper como fallback
            grupos_autores: Dict[str, List[str]] = {}
            for ced in cedulas_paper:
                for g in cedula_a_grupos.get(ced, []):
                    if g:  # solo grupos con nombre
                        grupos_autores.setdefault(g, []).append(ced)

            if not grupos_autores:
                continue

            info_ref = autores[0]  # metadata de referencia

            # ── Oportunidad cohesión: ≥2 autores del paper en el mismo grupo ──
            for g, ceds_en_g in grupos_autores.items():
                unicos = set(ceds_en_g)
                if len(unicos) >= 2 and g not in grupos_paper:
                    oport_cohesion.append({
                        "titulo_norm":           titulo_norm,
                        "grupo_beneficiario":    g,
                        "grupo_actual_registro": sorted(grupos_paper),
                        "cedulas_en_grupo":      sorted(unicos),
                        "año":                   info_ref["año"],
                        "categoria":             info_ref["cat"],
                        "tipo":                  "cohesion",
                    })
                    resumen.setdefault(g, {"cohesion": 0, "colaboracion": 0})
                    resumen[g]["cohesion"] += 1

            # ── Oportunidad colaboración: coautores externos al grupo ──────────
            ceds_set = set(cedulas_paper)
            for g, ceds_en_g in grupos_autores.items():
                ceds_en_g_set = set(ceds_en_g)
                # Autores del paper que NO pertenecen a este grupo
                externos = [
                    c for c in ceds_set - ceds_en_g_set
                    if g not in cedula_a_grupos.get(c, [])
                ]
                if externos and g not in grupos_paper:
                    oport_colaboracion.append({
                        "titulo_norm":           titulo_norm,
                        "grupo_beneficiario":    g,
                        "grupo_actual_registro": sorted(grupos_paper),
                        "cedulas_en_grupo":      sorted(ceds_en_g_set),
                        "cedulas_externas":      externos,
                        "año":                   info_ref["año"],
                        "categoria":             info_ref["cat"],
                        "tipo":                  "colaboracion",
                    })
                    resumen.setdefault(g, {"cohesion": 0, "colaboracion": 0})
                    resumen[g]["colaboracion"] += 1

        return {
            "autores_multigrupo":       multigrupo,
            "oportunidades_cohesion":   oport_cohesion,
            "oportunidades_colaboracion": oport_colaboracion,
            "resumen_por_grupo":        resumen,
        }

    def generar_dataframe(self) -> pd.DataFrame:
        """Devuelve las oportunidades como DataFrame listo para exportar."""
        # Si ya fue calculado externamente, reutilizar
        resultado = getattr(self, "_analizar_resultado", None) or self.analizar()
        filas = []
        for op in resultado["oportunidades_cohesion"] + resultado["oportunidades_colaboracion"]:
            filas.append({
                "Grupo beneficiario": op["grupo_beneficiario"],
                "Tipo oportunidad": op["tipo"].capitalize(),
                "Título (normalizado)": op["titulo_norm"][:120],
                "Año": op.get("año", ""),
                "Categoría": op.get("categoria", ""),
                "Cédulas en el grupo": "; ".join(op["cedulas_en_grupo"]),
                "Grupo(s) donde está registrado": "; ".join(op["grupo_actual_registro"]),
            })
        df = pd.DataFrame(filas)
        if not df.empty:
            df = df.sort_values(["Grupo beneficiario", "Tipo oportunidad"])
        return df
