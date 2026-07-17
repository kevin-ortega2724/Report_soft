"""
proyeccion_957.py

Proyecta cómo quedaría cada grupo (de los 75/125 con línea base oficial en
medicion_957.xlsx) si la Convocatoria 957 se recalculara HOY, usando su
producción de artículos ACTUAL (BD interna, alimentada por GrupLAC) cruzada
contra la clasificación real de la revista en Publindex (por ISSN + año).

Alcance de la proyección (documentado, no oculto -- ver
UI_clasificacion/data_sim/publindex.csv y obtener_publindex.py):
  - Solo se proyectan ARTÍCULOS (aportan a NC_TOP: ART_A1/ART_A2 -- y a
    NC_A: ART_B/ART_C/ART_D), que es lo único para lo que se puede
    determinar la calidad real vía Publindex (ISSN + año -> categoría).
    Libros, capítulos, extensión (ASC), divulgación (DPC) y formación de
    recurso humano (FRH_A/FRH_B) NO se proyectan -- quedan en su valor
    oficial de medicion_957.xlsx, sin cambios. El campo interno
    'categoria' de publicaciones no sirve para libros/capítulos (mezcla
    tipos descriptivos como "Libro de Texto" con códigos de incentivos
    internos UTP/CIARP como "10C", no la calidad real A1/A/B/C).
  - Solo aplica a los grupos con línea base oficial -- sin eso no hay
    `ratio` que calibrar (ver Simulador957.proyectar_productos).
  - Publindex (dataset abierto) no tiene vigencia 2023/2024 todavía (tope
    real: 2022) -- artículos publicados después usan la vigencia más
    reciente disponible para esa revista como aproximación documentada.
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from simulador_957 import Simulador957  # noqa: E402
from main_10 import DatabaseManager  # noqa: E402

PUBLINDEX_CSV = Path(__file__).parent / "data_sim" / "publindex.csv"

# Ventana de observación oficial de la Convocatoria 957: 1 ene 2019 - 31 dic
# 2023 (ver "Ficha de Asesoría..." en data_sim/). Solo lo publicado DESPUÉS
# de ese cierre cuenta como "nuevo" para la proyección.
FIN_VENTANA_OFICIAL = 2023

_SUBTIPO_POR_CATEGORIA_ARTICULO = {
    "A1": "ART_A1",
    "A2": "ART_A2",
    "B": "ART_B",
    "C": "ART_C",
    "D": "ART_D",
}

TOPE_VIGENCIA_PUBLINDEX = 2022


def _cargar_publindex(ruta: Path = PUBLINDEX_CSV) -> tuple[dict, dict]:
    """({(issn, anio): categoria}, {issn: anio_max_disponible})."""
    mapa: dict[tuple[str, int], str] = {}
    max_anio_por_issn: dict[str, int] = {}
    with open(ruta, encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            anio_raw = fila.get("anio_vigencia")
            categoria = (fila.get("categoria") or "").strip()
            if not anio_raw or not categoria:
                continue
            anio = int(float(anio_raw))
            for issn in (fila.get("issn_impreso"), fila.get("issn_linea")):
                issn = (issn or "").strip()
                if not issn:
                    continue
                mapa[(issn, anio)] = categoria
                if anio > max_anio_por_issn.get(issn, 0):
                    max_anio_por_issn[issn] = anio
    return mapa, max_anio_por_issn


class ProyeccionPublindex:
    def __init__(self, db, ruta_publindex: Path = PUBLINDEX_CSV):
        self.db = db
        self._mapa, self._max_anio_por_issn = _cargar_publindex(ruta_publindex)
        # Los nombres de grupo de medicion_957.xlsx casi nunca calzan
        # literal contra grupos.grupo de la BD interna (capitalización,
        # siglas, puntuación -- confirmado real: solo 6/67 calzan exacto).
        # Se reusa el mismo emparejamiento difuso que ya usa Simulador957
        # (_mejor_coincidencia: exacto normalizado -> contención -> Jaccard
        # >= 0.5) para resolver "Automática" -> "AUTOMÁTICA", etc.
        self._candidatos_bd = [
            r[0] for r in db.conn.cursor().execute(
                "SELECT DISTINCT grupo FROM grupos "
                "WHERE grupo IS NOT NULL AND grupo != '' "
                "AND grupo NOT LIKE '%SEMILLERO%' AND grupo NOT LIKE '%Semillero%' "
                "AND grupo NOT LIKE '%semillero%'"
            ).fetchall()
        ]
        self._cache_resolucion: dict[str, str | None] = {}

    def _nombre_en_bd(self, nombre_grupo: str) -> str | None:
        if nombre_grupo not in self._cache_resolucion:
            self._cache_resolucion[nombre_grupo] = Simulador957._mejor_coincidencia(
                nombre_grupo, self._candidatos_bd)
        return self._cache_resolucion[nombre_grupo]

    def categoria_de(self, issn: str, anio_publicacion: int) -> str | None:
        issn = (issn or "").strip()
        if not issn:
            return None
        anio_consulta = min(anio_publicacion, TOPE_VIGENCIA_PUBLINDEX)
        categoria = self._mapa.get((issn, anio_consulta))
        if categoria:
            return categoria
        max_anio = self._max_anio_por_issn.get(issn)
        if max_anio is not None:
            return self._mapa.get((issn, max_anio))
        return None

    def detalle_articulos_nuevos(self, grupo: str) -> list[dict]:
        """Detalle fila por fila (para trazabilidad en la UI) de los
        artículos publicados DESPUÉS de la ventana oficial, con su
        clasificación Publindex (o None si no se encontró la revista/ISSN).
        'grupo' es el nombre tal como aparece en medicion_957.xlsx."""
        nombre_bd = self._nombre_en_bd(grupo)
        if nombre_bd is None:
            return []

        cur = self.db.conn.cursor()
        filas = cur.execute('''
            SELECT p.titulo, p.issn_isbn, p.año FROM publicaciones p
            WHERE p.cedula IN (SELECT cedula FROM grupos WHERE grupo = ?)
              AND p.año > ?
              AND p.issn_isbn IS NOT NULL AND p.issn_isbn != ''
        ''', (nombre_bd, FIN_VENTANA_OFICIAL)).fetchall()

        detalle = []
        for titulo, issn, anio in filas:
            if not anio:
                continue
            categoria = self.categoria_de(issn, int(anio))
            detalle.append({
                "titulo": titulo,
                "issn": issn,
                "anio": int(anio),
                "categoria_publindex": categoria,
                "subtipo": _SUBTIPO_POR_CATEGORIA_ARTICULO.get(categoria),
            })
        return detalle

    def ajustes_articulos_nuevos(self, grupo: str) -> dict[str, int]:
        """{subtipo: n} agregados -- ver detalle_articulos_nuevos para el
        desglose fila por fila."""
        ajustes: dict[str, int] = {}
        for fila in self.detalle_articulos_nuevos(grupo):
            if fila["subtipo"]:
                ajustes[fila["subtipo"]] = ajustes.get(fila["subtipo"], 0) + 1
        return ajustes


def proyectar_todos(sim: Simulador957, proyector: ProyeccionPublindex) -> dict[str, dict]:
    """Corre la proyección para todos los grupos con línea base oficial.
    Devuelve {grupo: {oficial, proyectado (ResultadoSimulacion completos,
    con .valores_simulados/.indices_simulados/.condiciones/.detalle_ajustes),
    categoria_oficial, categoria_proyectada, ajustes, detalle_articulos,
    n_articulos_nuevos_clasificados}}."""
    resultados = {}
    for grupo in sim.df_grupos["nombre_grupo"].unique():
        detalle_articulos = proyector.detalle_articulos_nuevos(grupo)
        ajustes = {}
        for fila in detalle_articulos:
            if fila["subtipo"]:
                ajustes[fila["subtipo"]] = ajustes.get(fila["subtipo"], 0) + 1

        oficial = sim.simular(grupo, modo="verificacion")
        proyectado = sim.proyectar_productos(grupo, ajustes) if ajustes else oficial
        resultados[grupo] = {
            "oficial": oficial,
            "proyectado": proyectado,
            "categoria_oficial": oficial.categoria_simulada,
            "categoria_proyectada": proyectado.categoria_simulada,
            "ajustes": ajustes,
            "detalle_articulos": detalle_articulos,
            "n_articulos_nuevos_clasificados": sum(ajustes.values()),
        }
    return resultados


if __name__ == "__main__":
    db = DatabaseManager()
    sim = Simulador957()
    proyector = ProyeccionPublindex(db)
    resultados = proyectar_todos(sim, proyector)

    cambios = {g: r for g, r in resultados.items() if r["categoria_oficial"] != r["categoria_proyectada"]}
    print(f"{len(resultados)} grupos con línea base oficial evaluados.")
    print(f"{len(cambios)} con cambio de categoría proyectada.")
    for grupo, r in list(cambios.items())[:15]:
        print(f"  {grupo}: {r['categoria_oficial']} -> {r['categoria_proyectada']} "
              f"({r['n_articulos_nuevos_clasificados']} artículos nuevos clasificados, {r['ajustes']})")
