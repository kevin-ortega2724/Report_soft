"""
Funciones utilitarias compartidas entre los módulos de ReportSoft.
"""

import os
import re
import sys
from pathlib import Path

import pandas as pd
from unidecode import unidecode

from constants import CLASIF_957, VARIANTES_NOMBRES, HOJAS_NO_PRODUCTO, CATEGORIAS_PRINCIPALES


def obtener_directorio_base() -> Path:
    """Devuelve el directorio raíz del proyecto (donde está run.py, data/, src/)."""
    if getattr(sys, "frozen", False):
        return Path(os.path.dirname(sys.executable))
    return Path(os.path.dirname(os.path.abspath(__file__))).parent


def limpiar_cedula(cedula) -> str:
    """Devuelve solo dígitos de un valor que representa una cédula."""
    if pd.isna(cedula):
        return ""
    texto = str(cedula).strip().replace(".", "").replace(",", "").replace("-", "")
    return re.sub(r"[^\d]", "", texto)


def limpiar_texto(texto) -> str:
    """Devuelve cadena limpia o vacía si el valor es nulo/nan."""
    if pd.isna(texto) or str(texto).lower() in ("nan", "none", ""):
        return ""
    return str(texto).strip()


def normalizar_nombre(texto: str) -> str:
    """Normaliza un nombre para comparaciones robustas (sin tildes, minúsculas)."""
    if not texto:
        return ""
    resultado = unidecode(str(texto).lower()).strip()
    return re.sub(r"\s+", " ", resultado).strip(" ,.;:!?-_")


def norm_text(texto: str) -> str:
    """Alias de normalizar_nombre para uso en detección de columnas."""
    return normalizar_nombre(texto)


def normalizar_columna(nombre: str) -> str:
    """Normaliza el encabezado de una columna de Excel (minúsculas, sin tildes,
    cualquier carácter no alfanumérico → guion bajo) para compararlo de forma robusta."""
    return re.sub(r"[^a-z0-9]", "_", unidecode(str(nombre)).lower().strip())


def normalizar_nombre_hoja(nombre: str) -> str:
    """Normaliza el nombre de una hoja Excel para búsqueda flexible."""
    return re.sub(r"\s+", " ", unidecode(str(nombre).lower().strip()))


def limpiar_nombre_archivo(nombre: str, max_length: int = 150) -> str:
    """Sanitiza un texto para usarlo como nombre de archivo."""
    if not nombre:
        return "sin_nombre"
    reemplazos = {
        ":": "-", "/": "-", "\\": "-", "<": "_", ">": "_",
        '"': "", "'": "", "|": "-", "?": "", "*": "",
        "¿": "", "¡": "", "!": "", "@": "_at_", "#": "_num_",
        "$": "_", "%": "_pct_", "&": "_y_", "(": "_", ")": "_",
        "[": "_", "]": "_", "{": "_", "}": "_", "=": "_",
        "+": "_", "~": "_", "`": "", "^": "_", ";": "_",
        ",": "_", ".": "_",
    }
    resultado = str(nombre).strip()
    for char, rep in reemplazos.items():
        resultado = resultado.replace(char, rep)
    resultado = resultado.replace(" ", "_")
    resultado = re.sub(r"[-_]+", "_", resultado).strip("_-")
    if len(resultado) > max_length:
        resultado = resultado[:max_length].rstrip("_-")
    return resultado or "archivo"


def clasificar_producto_957(nombre_hoja: str):
    """
    Clasifica un nombre de hoja según CLASIF_957.

    Devuelve (categoria, subcategoria, nombre_oficial) o None si no coincide.
    Devuelve False si la hoja es de datos generales (no un producto).
    """
    nombre_norm = normalizar_nombre_hoja(nombre_hoja)

    # Hojas de datos generales → no son productos
    if nombre_norm in HOJAS_NO_PRODUCTO:
        return False
    for hoja_no in HOJAS_NO_PRODUCTO:
        if nombre_norm.startswith(hoja_no) or hoja_no.startswith(nombre_norm):
            return False

    # 1. Búsqueda exacta
    if nombre_hoja in CLASIF_957:
        cat, subcat = CLASIF_957[nombre_hoja]
        return (cat, subcat, nombre_hoja)

    # 2. Variantes normalizadas
    if nombre_norm in VARIANTES_NOMBRES:
        nombre_oficial = VARIANTES_NOMBRES[nombre_norm]
        if nombre_oficial in CLASIF_957:
            cat, subcat = CLASIF_957[nombre_oficial]
            return (cat, subcat, nombre_oficial)

    # 3. Coincidencia exacta normalizada contra cada clave del diccionario
    for nd, (cat, subcat) in CLASIF_957.items():
        if nombre_norm == normalizar_nombre_hoja(nd):
            return (cat, subcat, nd)

    # 4. Coincidencia parcial (contención)
    for nd, (cat, subcat) in CLASIF_957.items():
        nd_norm = normalizar_nombre_hoja(nd)
        if len(nombre_norm) > 15 and len(nd_norm) > 15:
            if nombre_norm in nd_norm or nd_norm in nombre_norm:
                return (cat, subcat, nd)

    # 5. Prefijo común (para nombres truncados por GrupLAC)
    if len(nombre_norm) >= 20:
        for nd, (cat, subcat) in CLASIF_957.items():
            nd_norm = normalizar_nombre_hoja(nd)
            if len(nd_norm) >= 20:
                pref = min(len(nombre_norm), len(nd_norm), 25)
                if nombre_norm[:pref] == nd_norm[:pref]:
                    return (cat, subcat, nd)

    return None


def get_orden_categoria(categoria: str) -> int:
    """Devuelve el orden numérico de una categoría principal."""
    for i, c in enumerate(CATEGORIAS_PRINCIPALES, 1):
        if categoria.startswith(str(i) + "."):
            return i
    return 99


def get_orden_producto(nombre: str) -> int:
    """Devuelve la posición de un producto en CLASIF_957."""
    keys = list(CLASIF_957.keys())
    return keys.index(nombre) if nombre in keys else 999


def tokenizar_autores(texto: str) -> set:
    """Extrae tokens de apellidos/nombres normalizados (mínimo 4 caracteres)."""
    if not texto or texto == "No especificado":
        return set()
    normalizado = unidecode(str(texto).lower())
    normalizado = re.sub(r"[^a-z\s]", " ", normalizado)
    stop = {
        "universidad", "tecnologica", "pereira", "autor", "autores",
        "nombre", "colombia", "grupo", "investigacion", "doctor",
        "magister", "phd", "msc", "prof", "email", "correo",
    }
    return {tok for tok in normalizado.split() if len(tok) >= 4} - stop
