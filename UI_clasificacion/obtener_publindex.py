"""
obtener_publindex.py

Descarga el dataset abierto oficial "Revistas Indexadas, Índice Nacional
Publindex" (datos.gov.co, Socrata, id `mwmn-inyg`) y lo guarda localmente en
UI_clasificacion/data_sim/publindex.csv.

Por qué: la clasificación de calidad de un artículo (A1/A2/B/C), que es el
dato de mayor peso en la fórmula de la Convocatoria 957, NO viene en el
scraping público de GrupLAC (el perfil de un grupo solo muestra
título/revista/ISSN/año, no la categoría Publindex de la revista). Ese dato
sí es público, pero en un dataset distinto: la clasificación homologada de
revistas de Publindex, publicada como dato abierto por el gobierno
colombiano.

Cobertura confirmada: 6276 filas, vigencias (nro_ano) de 2010 a 2022 -- no
hay 2023/2024 todavía en el dataset abierto. Para artículos publicados en
años posteriores a 2022, `proyeccion_957.py` usa la vigencia 2022 (la más
reciente disponible) como aproximación documentada, no oficial.

Uso:
    python UI_clasificacion/obtener_publindex.py
"""

from __future__ import annotations

import csv
from pathlib import Path

import requests

BASE_URL = "https://www.datos.gov.co/resource/mwmn-inyg.json"
PAGE_SIZE = 1000
SALIDA = Path(__file__).parent / "data_sim" / "publindex.csv"

# Columnas crudas del dataset -> nombres normalizados que usa proyeccion_957.py
COLUMNAS = {
    "txt_issn_p": "issn_impreso",
    "txt_issn_l": "issn_linea",
    "nme_revista_in": "nombre_revista",
    "nro_ano": "anio_vigencia",
    "id_clas_rev": "categoria",
}


def descargar_publindex() -> list[dict]:
    filas = []
    offset = 0
    while True:
        resp = requests.get(
            BASE_URL,
            params={"$limit": PAGE_SIZE, "$offset": offset, "$select": ",".join(COLUMNAS)},
            timeout=60,
        )
        resp.raise_for_status()
        lote = resp.json()
        if not lote:
            break
        filas.extend(lote)
        offset += PAGE_SIZE
        if len(lote) < PAGE_SIZE:
            break
    return filas


def guardar_csv(filas: list[dict], ruta: Path = SALIDA) -> None:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    columnas_salida = list(COLUMNAS.values())
    with open(ruta, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=columnas_salida)
        writer.writeheader()
        for fila in filas:
            writer.writerow({
                nombre_norm: fila.get(nombre_crudo, "")
                for nombre_crudo, nombre_norm in COLUMNAS.items()
            })


def main():
    print("Descargando dataset Publindex (datos.gov.co)...")
    filas = descargar_publindex()
    guardar_csv(filas)
    print(f"{len(filas)} filas guardadas en {SALIDA}")


if __name__ == "__main__":
    main()
