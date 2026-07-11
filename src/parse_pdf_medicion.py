"""
parse_pdf_medicion.py

Extrae datos del "Proceso de Medición" GrupLAC (Convocatoria 957/2024).
Cada PDF contiene N grupos de una misma categoría.

Uso rápido:
    from parse_pdf_medicion import parse_and_export
    dfs = parse_and_export()          # lee data/pdf/, exporta a data/output/
    dfs = parse_and_export(out=None)  # solo devuelve DataFrames sin exportar

Retorna dict con 4 DataFrames:
    grupos      – una fila por grupo (nombre, director, categoría, indicador, área)
    productos   – una fila por (grupo, subtipo): total, ventana, lambda
    indicadores – una fila por (grupo, indicador): valor, valor_maximo, indice, ponderacion
    cuartiles   – una fila por (grupo, cuartil): min, q4, q3, q2, max, valor_grupo

Dependencias: pdfplumber  (pip install pdfplumber)
"""

from __future__ import annotations

import re
import pathlib
from typing import Optional
import pandas as pd

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

PDF_CATEGORIAS: dict[str, str] = {
    "A1s.pdf": "A1",
    "As.pdf":  "A",
    "Bs.pdf":  "B",
    "Cs.pdf":  "C",
    "NRs.pdf": "NR",
    "SRs.pdf": "SR",
}

# Nombre corto para cada sección de conteo de productos
_SECCION_MAP: list[tuple[str, str]] = [
    (r"nuevo conocimiento\s*[-–]\s*top",                        "NC_TOP"),
    (r"nuevo conocimiento\s+tipo\s+a\b",                        "NC_A"),
    (r"nuevo conocimiento\s+tipo\s+b\b",                        "NC_B"),
    (r"apropiaci[oó]n social del conocimiento",                  "ASC"),
    (r"divulgaci[oó]n p[uú]blica de la ciencia",                 "DPC"),
    (r"formaci[oó]n de recurso humano.*tipo\s+a\b",              "FRH_A"),
    (r"formaci[oó]n de recurso humano.*tipo\s+b\b",              "FRH_B"),
    (r"formaci[oó]n.*tipo\s+a\b",                               "FRH_A"),
    (r"formaci[oó]n.*tipo\s+b\b",                               "FRH_B"),
]

# Indicadores conocidos  (re.search sobre el texto ya sin prefijos "productos de"/"índice de")
_INDICADOR_MAP: list[tuple[str, str]] = [
    (r"nuevo conocimiento top",                       "NC_TOP"),
    (r"nuevo conocimiento tipo a\b",                  "NC_A"),
    (r"nuevo conocimiento tipo b\b",                  "NC_B"),
    (r"apropiaci[oó]n social",                        "ASC"),
    (r"divulgaci[oó]n p[uú]blica",                    "DPC"),
    (r"formaci[oó]n tipo a\b",                        "FRH_A"),
    (r"formaci[oó]n tipo b\b",                        "FRH_B"),
    (r"cohesi[oó]n",                                  "cohesion"),
    (r"colaboraci[oó]n",                              "colaboracion"),
]

_FLOAT = r"[\d]+\.[\d]+"  # patron numero float


# ---------------------------------------------------------------------------
# Función principal
# ---------------------------------------------------------------------------

def parse_and_export(
    pdf_dir: str | pathlib.Path | None = None,
    out: str | pathlib.Path | None = "auto",
) -> dict[str, pd.DataFrame]:
    """
    Parsea todos los PDFs y exporta a Excel.

    Args:
        pdf_dir: Carpeta con los PDFs. Por defecto data/pdf/ relativa al proyecto.
        out: Ruta del Excel de salida.  "auto" → data/output/medicion_957.xlsx
             None → no exporta, solo devuelve DataFrames.

    Returns:
        {"grupos": df, "productos": df, "indicadores": df, "cuartiles": df}
    """
    import pdfplumber  # importación diferida para no romper si no está instalado

    base = pathlib.Path(__file__).parent.parent
    if pdf_dir is None:
        pdf_dir = base / "data" / "pdf"
    pdf_dir = pathlib.Path(pdf_dir)

    rows_grupos: list[dict] = []
    rows_productos: list[dict] = []
    rows_indicadores: list[dict] = []
    rows_cuartiles: list[dict] = []

    for pdf_name, cat in PDF_CATEGORIAS.items():
        path = pdf_dir / pdf_name
        if not path.exists():
            print(f"  [omitido] {path} no encontrado")
            continue
        print(f"Procesando {pdf_name} ({cat})…", end=" ", flush=True)
        grupos = _parse_pdf(path, cat, pdfplumber)
        print(f"{len(grupos)} grupos")

        for g in grupos:
            nombre = g["nombre_grupo"]
            rows_grupos.append({
                k: v for k, v in g.items()
                if k not in ("productos", "indicadores", "cuartiles")
            })
            for p in g.get("productos", []):
                rows_productos.append({"grupo": nombre, "categoria": cat, **p})
            for ind in g.get("indicadores", []):
                rows_indicadores.append({"grupo": nombre, "categoria": cat, **ind})
            for c in g.get("cuartiles", []):
                rows_cuartiles.append({"grupo": nombre, "categoria": cat, **c})

    dfs = {
        "grupos":      pd.DataFrame(rows_grupos),
        "productos":   pd.DataFrame(rows_productos),
        "indicadores": pd.DataFrame(rows_indicadores),
        "cuartiles":   pd.DataFrame(rows_cuartiles),
    }

    if out == "auto":
        out = base / "data" / "output" / "medicion_957.xlsx"
    if out is not None:
        out = pathlib.Path(out)
        out.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            for sheet, df in dfs.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
        print(f"\nExportado -> {out}")

    return dfs


# ---------------------------------------------------------------------------
# Parseo interno
# ---------------------------------------------------------------------------

def _parse_pdf(path: pathlib.Path, categoria: str, pdfplumber) -> list[dict]:
    """Lee un PDF y devuelve lista de dicts, uno por grupo."""
    pages: list[str] = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            pages.append(page.extract_text(x_tolerance=2) or "")

    # Segmentar por grupo (cada portada tiene "Director del Grupo")
    starts = [i for i, t in enumerate(pages) if "Director del Grupo" in t]
    groups = []
    for j, s in enumerate(starts):
        e = starts[j + 1] if j + 1 < len(starts) else len(pages)
        text = "\n".join(pages[s:e])
        try:
            g = _parse_group(text, categoria)
            groups.append(g)
        except Exception as ex:
            print(f"\n    [error] grupo {j+1} de {path.name}: {ex}")
    return groups


def _parse_group(text: str, categoria: str) -> dict:
    """Extrae todos los datos de un grupo desde su texto completo."""
    return {
        "categoria": categoria,
        **_portada(text),
        **_resumen(text),
        "productos":   _productos(text),
        "indicadores": _indicadores(text),
        "cuartiles":   _cuartiles(text),
    }


# ── Portada ────────────────────────────────────────────────────────────────

def _portada(text: str) -> dict:
    director = _match1(r"Director del Grupo:\s*(.+)", text) or ""
    nombre   = _match1(r"Nombre del Grupo:\s*(.+)",   text) or ""
    return {
        "nombre_grupo": nombre.strip(),
        "director":     director.strip(),
    }


# ── Resumen (categoría + indicador + área) ─────────────────────────────────

def _resumen(text: str) -> dict:
    cat = _match1(r"La categor[ií]a alcanzada por el grupo fue:\s*(\w+)", text) or ""
    ind = None
    m = re.search(
        r"El indicador para el grupo .+? es:\s*\n?\s*([\d.]+)",
        text,
    )
    if m:
        try:
            ind = float(m.group(1).strip())
        except ValueError:
            pass

    # Preferir la línea "6.1 Cuartiles ... para el área de conocimiento X"
    area = (
        _match1(
            r"6\.1\s+Cuartiles\s+.+?[aá]rea de conocimiento\s+([^\n]+)",
            text, flags=re.IGNORECASE,
        )
        or _match1(
            r"para el [aá]rea de conocimiento\s+([A-ZÁÉÍÓÚÑ][^\n]+)",
            text, flags=re.IGNORECASE,
        )
        or ""
    )
    area = re.sub(r"\s+del Grupo de Investigación.*", "", area).strip()

    return {
        "categoria_confirmada": cat.strip(),
        "indicador_grupo":      ind,
        "area_conocimiento":    area,
    }


# ── Conteo de productos (Paso 2) ──────────────────────────────────────────

def _productos(text: str) -> list[dict]:
    results: list[dict] = []
    seccion = "?"
    lines = text.split("\n")
    for line in lines:
        low = line.lower().strip()
        # Detectar encabezado de sección
        if "conteo de productos resultado de actividades de" in low:
            tail = re.sub(
                r".*conteo de productos resultado de actividades de\s*", "", low
            )
            seccion = _seccion(tail)
            continue

        # Detectar fila de datos: CÓDIGO int int float
        m = re.match(
            r"^([A-Z][A-Z0-9_]+)\s+(\d+)\s+(\d+)\s+(" + _FLOAT + r")\s*$",
            line.strip(),
        )
        if m:
            results.append({
                "seccion":   seccion,
                "subtipo":   m.group(1),
                "total":     int(m.group(2)),
                "ventana":   int(m.group(3)),
                "lambda_val": float(m.group(4)),
            })
    return results


def _seccion(tail: str) -> str:
    for pat, name in _SECCION_MAP:
        if re.search(pat, tail, re.IGNORECASE):
            return name
    return tail[:40]


# ── Valores de indicadores y tabla de índices (Pasos 3–5) ─────────────────

def _indicadores(text: str) -> list[dict]:
    """
    Extrae la tabla de indicadores (Paso 3) y la tabla de índices (Paso 4)
    y la tabla de ponderación (Paso 5) como filas combinadas.
    """
    # ── Paso 3: indicadores ────────────────────────────────────────────────
    ind_vals: dict[str, float] = {}
    # Buscar el bloque entre el encabezado y "Indicador de cohesión" o "Paso 4"
    block_match = re.search(
        r"Indicador\s+Valor del indicador obtenido por el grupo\s*\n"
        r"(.*?)"
        r"(?:Indicador de cohesi[oó]n|Paso 4)",
        text, re.DOTALL | re.IGNORECASE,
    )
    if block_match:
        block = block_match.group(1)
        for line in block.split("\n"):
            line = line.strip()
            if not line:
                continue
            # Línea: "NOMBRE ... VALOR"
            m = re.match(r"^(.+?)\s+(" + _FLOAT + r")\s*$", line)
            if m:
                clave = _clave_indicador(m.group(1).strip())
                if clave:
                    ind_vals[clave] = float(m.group(2))

    # ── Paso 4: índices ────────────────────────────────────────────────────
    idx_vals: dict[str, tuple[float, float, float]] = {}
    # Bloque entre "Paso 4" y "Paso 5"
    idx_block = re.search(
        r"Paso 4[:\s]+.+?\n(.*?)(?:Paso 5)",
        text, re.DOTALL | re.IGNORECASE,
    )
    if idx_block:
        _parse_three_float_table(idx_block.group(1), idx_vals)

    # ── Paso 5: ponderación ────────────────────────────────────────────────
    pond_vals: dict[str, tuple[float, float, float]] = {}
    pond_block = re.search(
        r"Paso 5[:\s]+.+?\n(.*?)(?:Paso 6|El indicador para el grupo)",
        text, re.DOTALL | re.IGNORECASE,
    )
    if pond_block:
        _parse_three_float_table(pond_block.group(1), pond_vals)

    # ── Combinar en lista de filas ─────────────────────────────────────────
    rows = []
    todas_claves = set(ind_vals) | set(idx_vals) | set(pond_vals)
    for clave in todas_claves:
        row: dict = {"indicador": clave}
        row["valor_indicador"] = ind_vals.get(clave)
        if clave in idx_vals:
            row["valor_maximo"], row["valor_grupo_idx"], row["valor_indice"] = idx_vals[clave]
        else:
            row["valor_maximo"] = row["valor_grupo_idx"] = row["valor_indice"] = None
        if clave in pond_vals:
            row["ponderacion"], row["indice_pond"], row["valor_ponderado"] = pond_vals[clave]
        else:
            row["ponderacion"] = row["indice_pond"] = row["valor_ponderado"] = None
        rows.append(row)
    return rows


def _parse_three_float_table(block: str, out: dict) -> None:
    """
    Intenta extraer filas "ETIQUETA float float float" de un bloque de texto.
    Maneja tanto filas en una sola línea como etiquetas que se parten en 2 líneas.
    """
    lines = [l.strip() for l in block.split("\n") if l.strip()]
    i = 0
    while i < len(lines):
        line = lines[i]
        # Caso 1: toda la fila en una línea → "ETIQ f f f"
        m = re.match(
            r"^(.+?)\s+(" + _FLOAT + r")\s+(" + _FLOAT + r")\s+(" + _FLOAT + r")\s*$",
            line,
        )
        if m:
            clave = _clave_indicador(m.group(1))
            if clave:
                out[clave] = (float(m.group(2)), float(m.group(3)), float(m.group(4)))
            i += 1
            continue

        # Caso 2: línea contiene solo 3 floats (etiqueta partida, números en el medio)
        m2 = re.match(
            r"^(" + _FLOAT + r")\s+(" + _FLOAT + r")\s+(" + _FLOAT + r")\s*$",
            line,
        )
        if m2:
            # Buscar la etiqueta buscando hacia adelante (siguiente línea con texto)
            label = ""
            if i + 1 < len(lines) and not re.match(r"^" + _FLOAT, lines[i + 1]):
                label = lines[i - 1] if i > 0 else ""
                # la parte final del nombre está en lines[i+1]
                label = (label + " " + lines[i + 1]).strip()
            elif i > 0:
                label = lines[i - 1]
            clave = _clave_indicador(label)
            if clave:
                out[clave] = (float(m2.group(1)), float(m2.group(2)), float(m2.group(3)))
            i += 1
            continue

        i += 1


def _clave_indicador(texto: str) -> Optional[str]:
    """Convierte un nombre de indicador en su clave corta."""
    texto_low = texto.lower().strip()
    # Limpiar prefijos comunes de la tabla de ponderación
    texto_low = re.sub(r"^[íi]ndice de\s+", "", texto_low)
    texto_low = re.sub(r"^productos de\s+", "", texto_low)
    for pat, clave in _INDICADOR_MAP:
        if re.search(pat, texto_low):
            return clave
    return None


# ── Cuartiles (Paso 6) ─────────────────────────────────────────────────────

def _cuartiles(text: str) -> list[dict]:
    """
    Extrae los bloques de cuartiles.
    Formato:
        Cuartil de NOMBRE
        Mínimo Cuartil 4 Cuartil 3 Cuartil 2 Máximo
        n1 n2 n3 n4 n5
        El valor del indicador para su grupo de investigación es: n6
    """
    pattern = re.compile(
        r"Cuartil(?:es)?\s+(?:de\s+|para\s+el\s+)?(.+?)\n"
        r"M[íi]nimo\s+Cuartil\s+4\s+Cuartil\s+3\s+Cuartil\s+2\s+M[áa]ximo\n"
        r"(" + _FLOAT + r")\s+(" + _FLOAT + r")\s+(" + _FLOAT + r")\s+"
        r"(" + _FLOAT + r")\s+(" + _FLOAT + r")\n"
        r"El valor del indicador para su grupo de investigaci[oó]n es:\s*(" + _FLOAT + r")",
        re.IGNORECASE,
    )
    rows = []
    for m in pattern.finditer(text):
        rows.append({
            "cuartil":    m.group(1).strip(),
            "min":        float(m.group(2)),
            "q4":         float(m.group(3)),
            "q3":         float(m.group(4)),
            "q2":         float(m.group(5)),
            "max":        float(m.group(6)),
            "valor_grupo": float(m.group(7)),
        })
    return rows


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def _match1(pattern: str, text: str, flags: int = 0) -> Optional[str]:
    m = re.search(pattern, text, flags)
    return m.group(1) if m else None


# ---------------------------------------------------------------------------
# BenchmarksMedicion957 – acceso a datos reales para simulación
# ---------------------------------------------------------------------------

# Nombres de cuartiles → clave interna de indicador
_CUARTIL_A_IND: list[tuple[str, str]] = [
    (r"nuevo conocimiento top",   "NC_TOP"),
    (r"nuevo conocimiento a\b",   "NC_A"),
    (r"nuevo conocimiento b\b",   "NC_B"),
    (r"apropiaci[oó]n social",    "ASC"),
    (r"divulgaci[oó]n p[uú]blica", "DPC"),
    (r"formaci[oó]n.*\ba\b",      "FRH_A"),
    (r"formaci[oó]n.*\bb\b",      "FRH_B"),
    (r"cohesi[oó]n",              "cohesion"),
    (r"colaboraci[oó]n",          "colaboracion"),
    (r"[aá]rea de conocimiento|grupo de investigaci[oó]n", "IG"),
]

# Ponderaciones fijas de la convocatoria 957 (por si el Excel no las trae)
PONDERACIONES_957: dict[str, float] = {
    "NC_TOP":      3.7,
    "NC_A":        2.3,
    "NC_B":        0.4,
    "ASC":         1.5,
    "DPC":         0.5,
    "FRH_A":       1.0,
    "FRH_B":       0.2,
    "cohesion":    0.1,
    "colaboracion": 0.3,
}


class BenchmarksMedicion957:
    """
    Carga los datos reales de la medición 957 (Excel generado por parse_and_export)
    y provee métodos para usarlos en la simulación de categorías.

    Uso típico:
        bm = BenchmarksMedicion957()
        areas = bm.areas
        cuartiles = bm.cuartiles_ig(area)   # umbrales del Indicador de Grupo
        maximos  = bm.maximos_ind(area)     # max de cada sub-indicador en el área
        ponds    = bm.ponderaciones         # pesos de cada sub-indicador
    """

    def __init__(self, excel_path: str | pathlib.Path | None = None):
        if excel_path is None:
            excel_path = (
                pathlib.Path(__file__).parent.parent
                / "data" / "output" / "medicion_957.xlsx"
            )
        self._path = pathlib.Path(excel_path)
        self._df_grupos:      pd.DataFrame = pd.DataFrame()
        self._df_cuartiles:   pd.DataFrame = pd.DataFrame()
        self._df_indicadores: pd.DataFrame = pd.DataFrame()
        self._ok = False
        self._load()

    def _load(self) -> None:
        if not self._path.exists():
            return
        try:
            self._df_grupos      = pd.read_excel(self._path, sheet_name="grupos")
            self._df_cuartiles   = pd.read_excel(self._path, sheet_name="cuartiles")
            self._df_indicadores = pd.read_excel(self._path, sheet_name="indicadores")
            self._ok = True
        except Exception:
            pass

    @property
    def disponible(self) -> bool:
        return self._ok

    @property
    def areas(self) -> list[str]:
        if self._df_grupos.empty:
            return []
        return sorted(
            self._df_grupos["area_conocimiento"].dropna().unique().tolist()
        )

    @property
    def ponderaciones(self) -> dict[str, float]:
        """Retorna {indicador: ponderación} — usa valores del Excel si están presentes."""
        result = dict(PONDERACIONES_957)
        if self._df_indicadores.empty:
            return result
        df = self._df_indicadores.dropna(subset=["ponderacion"])
        for _, row in df.iterrows():
            ind = row.get("indicador", "")
            pond = row.get("ponderacion")
            if ind and pd.notna(pond) and ind not in result:
                result[ind] = float(pond)
        return result

    def grupos_en_area(self, area: str) -> list[str]:
        """Nombres de los grupos reales en el área dada."""
        mask = self._df_grupos["area_conocimiento"].str.contains(
            re.escape(area), case=False, na=False
        )
        return self._df_grupos.loc[mask, "nombre_grupo"].tolist()

    def cuartiles_ig(self, area: str) -> dict | None:
        """
        Cuartiles del Indicador de Grupo para el área dada.
        Retorna {'min', 'q4', 'q3', 'q2', 'max'} o None si no hay datos.
        """
        grupos = set(self.grupos_en_area(area))
        if not grupos:
            return None
        mask_g = self._df_cuartiles["grupo"].isin(grupos)
        mask_c = self._df_cuartiles["cuartil"].str.contains(
            r"[aá]rea de conocimiento|Grupo de Investigaci[oó]n",
            case=False, na=False, regex=True,
        )
        df = self._df_cuartiles[mask_g & mask_c]
        if df.empty:
            return None
        # Las filas del mismo área deberían tener los mismos valores → tomamos la mediana
        return {
            "min": float(df["min"].median()),
            "q4":  float(df["q4"].median()),
            "q3":  float(df["q3"].median()),
            "q2":  float(df["q2"].median()),
            "max": float(df["max"].median()),
        }

    def maximos_ind(self, area: str) -> dict[str, float]:
        """
        Valor máximo de cada sub-indicador para el área dada.
        Usado como denominador al calcular el índice: indice = valor / maximo.
        Retorna {clave_ind: max_value}.
        """
        grupos = set(self.grupos_en_area(area))
        if not grupos:
            return {}
        df = self._df_cuartiles[self._df_cuartiles["grupo"].isin(grupos)]
        result: dict[str, float] = {}
        for _, row in df.iterrows():
            cuartil_name = str(row.get("cuartil", "")).lower()
            max_val = row.get("max")
            if not pd.notna(max_val) or max_val == 0:
                continue
            for pat, clave in _CUARTIL_A_IND:
                if re.search(pat, cuartil_name, re.IGNORECASE):
                    # Tomamos el máximo de los máximos (por si difiere entre grupos)
                    if clave not in result or float(max_val) > result[clave]:
                        result[clave] = float(max_val)
                    break
        return result

    def cuartil_de_valor(
        self, valor: float, cuartiles: dict
    ) -> tuple[int, str]:
        """
        Dado un valor y los cuartiles {min,q4,q3,q2,max}, determina en qué cuartil
        está el valor.
        Retorna (n_cuartil, descripción) donde n_cuartil=1 es el más alto.
        """
        if valor >= cuartiles.get("q2", float("inf")):
            return 1, "Cuartil 1 (25% superior)"
        if valor >= cuartiles.get("q3", float("inf")):
            return 2, "Cuartil 2 (50-75%)"
        if valor >= cuartiles.get("q4", float("inf")):
            return 3, "Cuartil 3 (25-50%)"
        return 4, "Cuartil 4 (25% inferior)"

    def calcular_ig(
        self,
        sub_indicadores: dict[str, float],
        area: str,
    ) -> dict:
        """
        Calcula el Indicador de Grupo completo a partir de sub-indicadores internos
        usando los máximos y ponderaciones reales de MinCiencias.

        Args:
            sub_indicadores: {clave_ind: valor}  (ej. NC_TOP, NC_A, …)
            area: área de conocimiento del grupo

        Returns:
            {
              'indicador_grupo': float,
              'detalle': {clave: {valor, maximo, indice, ponderacion, aporte}},
              'cuartil_ig': int,
              'desc_cuartil': str,
            }
        """
        maximos   = self.maximos_ind(area)
        ponds     = self.ponderaciones
        cuartiles = self.cuartiles_ig(area)

        ig = 0.0
        detalle: dict[str, dict] = {}

        for clave, pond in PONDERACIONES_957.items():
            valor  = sub_indicadores.get(clave, 0.0)
            maximo = maximos.get(clave, 0.0)
            indice = (valor / maximo) if maximo > 0 else 0.0
            aporte = indice * pond
            ig    += aporte
            detalle[clave] = {
                "valor":      round(valor, 4),
                "maximo":     round(maximo, 4),
                "indice":     round(indice, 6),
                "ponderacion": pond,
                "aporte":     round(aporte, 6),
            }

        cuartil_n, desc = (
            self.cuartil_de_valor(ig, cuartiles)
            if cuartiles else (4, "sin datos de área")
        )

        return {
            "indicador_grupo": round(ig, 6),
            "detalle":         detalle,
            "cuartil_ig":      cuartil_n,
            "desc_cuartil":    desc,
            "cuartiles_ref":   cuartiles,
        }

    def condiciones_categoria(
        self,
        sub_indicadores: dict[str, float],
        ig_result: dict,
        area: str,
    ) -> dict[str, dict]:
        """
        Evalúa si el grupo cumple las condiciones de cada categoría A1/A/B/C.
        Retorna {cat: {'cumple': bool, 'condiciones': [(desc, bool)]}}
        """
        ig      = ig_result["indicador_grupo"]
        cuarts  = ig_result.get("cuartiles_ref") or {}
        nc_top  = sub_indicadores.get("NC_TOP", 0)
        nc_a    = sub_indicadores.get("NC_A", 0)
        asc     = sub_indicadores.get("ASC", 0)
        dpc     = sub_indicadores.get("DPC", 0)
        frh_a   = sub_indicadores.get("FRH_A", 0)
        cohesion = sub_indicadores.get("cohesion", 0)

        cuartil_ig,  _ = self.cuartil_de_valor(ig, cuarts) if cuarts else (4, "")
        q2_top = (self.maximos_ind(area).get("NC_TOP", 0)
                  * self.cuartiles_ig(area).get("q2", 0)
                  if self.cuartiles_ig(area) else 0)

        # Cuartil del NC_TOP
        cuartiles_top_ref = None
        grupos = set(self.grupos_en_area(area))
        df_top = self._df_cuartiles[
            self._df_cuartiles["grupo"].isin(grupos)
            & self._df_cuartiles["cuartil"].str.contains(
                "Nuevo Conocimiento TOP", case=False, na=False)
        ]
        if not df_top.empty:
            cuartiles_top_ref = {
                "min": float(df_top["min"].median()),
                "q4":  float(df_top["q4"].median()),
                "q3":  float(df_top["q3"].median()),
                "q2":  float(df_top["q2"].median()),
                "max": float(df_top["max"].median()),
            }
        cuartil_top, _ = (
            self.cuartil_de_valor(nc_top, cuartiles_top_ref)
            if cuartiles_top_ref else (4, "")
        )

        cats: dict[str, dict] = {}

        # ── A1 ──────────────────────────────────────────────────────────────
        conds_a1 = [
            (f"IG en cuartil 1 (IG={ig:.4f} ≥ Q2={cuarts.get('q2',0):.4f})",
             cuartil_ig <= 1),
            (f"NC_TOP en cuartil 1 (NC_TOP={nc_top:.3f})",
             cuartil_top <= 1),
            (f"ASC > 0 o DPC > 0 (ASC={asc:.3f}, DPC={dpc:.3f})",
             asc > 0 or dpc > 0),
            (f"FRH_A > 0 (FRH_A={frh_a:.3f})", frh_a > 0),
            (f"Cohesión > 0 (cohesion={cohesion:.3f})", cohesion > 0),
        ]
        cats["A1"] = {
            "cumple": all(c[1] for c in conds_a1),
            "condiciones": conds_a1,
        }

        # ── A ───────────────────────────────────────────────────────────────
        conds_a = [
            (f"IG en cuartil 1 o 2 (IG={ig:.4f} ≥ Q3={cuarts.get('q3',0):.4f})",
             cuartil_ig <= 2),
            (f"NC_TOP > 0 o NC_A > 0 o DPC > 0",
             nc_top > 0 or nc_a > 0 or dpc > 0),
            (f"ASC > 0 (ASC={asc:.3f})", asc > 0),
            (f"FRH_A > 0 (FRH_A={frh_a:.3f})", frh_a > 0),
            (f"Cohesión > 0 (cohesion={cohesion:.3f})", cohesion > 0),
        ]
        cats["A"] = {
            "cumple": all(c[1] for c in conds_a),
            "condiciones": conds_a,
        }

        # ── B ───────────────────────────────────────────────────────────────
        conds_b = [
            (f"IG en cuartil 1-3 (IG={ig:.4f} ≥ Q4={cuarts.get('q4',0):.4f})",
             cuartil_ig <= 3),
            (f"NC_TOP > 0 o NC_A > 0 o NC_B > 0",
             nc_top > 0 or nc_a > 0 or sub_indicadores.get("NC_B", 0) > 0),
            (f"FRH_A > 0 (FRH_A={frh_a:.3f})", frh_a > 0),
        ]
        cats["B"] = {
            "cumple": all(c[1] for c in conds_b),
            "condiciones": conds_b,
        }

        # ── C ───────────────────────────────────────────────────────────────
        conds_c = [
            (f"IG > mínimo (IG={ig:.4f} > 0)", ig > 0),
            (f"NC_TOP > 0 o NC_A > 0 o NC_B > 0 o DPC > 0 o FRH_A > 0",
             any(sub_indicadores.get(k, 0) > 0
                 for k in ["NC_TOP", "NC_A", "NC_B", "DPC", "FRH_A"])),
        ]
        cats["C"] = {
            "cumple": all(c[1] for c in conds_c),
            "condiciones": conds_c,
        }

        return cats

    def estadisticas_por_categoria(
        self, area: str, categorias: list[str] | None = None
    ) -> dict[str, dict[str, dict]]:
        """
        Para cada categoría solicitada, devuelve estadísticas reales de los
        sub-indicadores de los grupos de esa categoría en el área dada.

        Returns:
            {
              "A1": {
                "NC_TOP": {"min": X, "median": Y, "max": Z, "n": N},
                "NC_A":   {...},
                ...
                "IG":     {"min": X, "median": Y, "max": Z, "n": N},
              },
              "A": {...},
              ...
            }
        """
        if categorias is None:
            categorias = ["A1", "A", "B", "C"]

        # Grupos de esa área por categoría
        df_g = self._df_grupos
        df_c = self._df_cuartiles
        df_i = self._df_indicadores

        resultado: dict[str, dict] = {}

        for cat in categorias:
            mask = (
                df_g["area_conocimiento"].str.contains(re.escape(area), case=False, na=False)
                & (df_g["categoria"] == cat)
            )
            grupos_cat = set(df_g.loc[mask, "nombre_grupo"].tolist())
            if not grupos_cat:
                resultado[cat] = {}
                continue

            stats: dict[str, dict] = {}

            # — Sub-indicadores desde df_indicadores —
            df_ind_cat = df_i[df_i["grupo"].isin(grupos_cat)]
            for clave_bm in list(_CUARTIL_A_IND_KEYS) + ["cohesion", "colaboracion"]:
                vals = df_ind_cat.loc[
                    df_ind_cat["indicador"] == clave_bm, "valor_indicador"
                ].dropna().tolist()
                if vals:
                    import statistics
                    stats[clave_bm] = {
                        "min":    round(min(vals), 3),
                        "median": round(statistics.median(vals), 3),
                        "max":    round(max(vals), 3),
                        "n":      len(vals),
                    }

            # — Indicador de Grupo desde df_grupos —
            ig_vals = df_g.loc[
                df_g["nombre_grupo"].isin(grupos_cat), "indicador_grupo"
            ].dropna().tolist()
            if ig_vals:
                import statistics
                stats["IG"] = {
                    "min":    round(min(ig_vals), 4),
                    "median": round(statistics.median(ig_vals), 4),
                    "max":    round(max(ig_vals), 4),
                    "n":      len(ig_vals),
                }

            resultado[cat] = stats

        return resultado

    def brechas_para_a1(
        self,
        sub_indicadores: dict[str, float],
        ig_calculado: float,
        area: str,
    ) -> list[dict]:
        """
        Retorna una lista de brechas concretas que el grupo debe cerrar para A1.
        Cada brecha incluye la condición, el valor actual, el objetivo y la
        referencia (mínimo de grupos A1 reales del área).

        Returns lista de dicts:
            {
              "condicion": str,
              "cumple": bool,
              "actual": float,
              "objetivo_min": float,   # mínimo que tienen los grupos A1 reales
              "objetivo_q1": float,    # umbral Q1 del cuartil oficial
              "brecha": float,
              "unidades_necesarias": str,  # estimación práctica
            }
        """
        cuarts_ig    = self.cuartiles_ig(area) or {}
        cuarts_top   = self._cuartiles_para_indicador(area, "Nuevo Conocimiento TOP")
        stats_a1     = self.estadisticas_por_categoria(area, ["A1"]).get("A1", {})
        maximos      = self.maximos_ind(area)

        brechas = []

        # ── 1. IG en Q1 ────────────────────────────────────────────────────
        obj_q1_ig  = cuarts_ig.get("q2", 0)
        obj_min_a1 = stats_a1.get("IG", {}).get("min", obj_q1_ig)
        cumple_ig  = ig_calculado >= obj_q1_ig
        brechas.append({
            "condicion":        "Indicador de Grupo en Q1 (top 25%)",
            "cumple":           cumple_ig,
            "actual":           round(ig_calculado, 4),
            "objetivo_q1":      round(obj_q1_ig, 4),
            "objetivo_min_a1":  round(obj_min_a1, 4),
            "brecha":           round(max(0, obj_q1_ig - ig_calculado), 4),
            "unidades_necesarias": (
                "" if cumple_ig else
                self._estimar_articulos_para_ig(
                    obj_q1_ig - ig_calculado, maximos
                )
            ),
        })

        # ── 2. NC_TOP en Q1 ────────────────────────────────────────────────
        nc_top_actual = sub_indicadores.get("NC_TOP", 0)
        obj_q1_top    = cuarts_top.get("q2", 0) if cuarts_top else 0
        obj_min_top   = stats_a1.get("NC_TOP", {}).get("min", obj_q1_top)
        cumple_top    = nc_top_actual >= obj_q1_top
        brecha_top    = max(0, obj_q1_top - nc_top_actual)
        # ¿Cuántos artículos TOP (lambda=1.0) para cerrar la brecha?
        arts_top = f"≈{int(brecha_top / 1.0) + 1} artículos Q1/Q2" if not cumple_top and brecha_top > 0 else ""
        brechas.append({
            "condicion":        "NC_TOP en Q1 (top 25% del área)",
            "cumple":           cumple_top,
            "actual":           round(nc_top_actual, 3),
            "objetivo_q1":      round(obj_q1_top, 3),
            "objetivo_min_a1":  round(obj_min_top, 3),
            "brecha":           round(brecha_top, 3),
            "unidades_necesarias": arts_top,
        })

        # ── 3. ASC > 0 o DPC > 0 ──────────────────────────────────────────
        asc = sub_indicadores.get("ASC", 0)
        dpc = sub_indicadores.get("DPC", 0)
        cumple_asc_dpc = asc > 0 or dpc > 0
        brechas.append({
            "condicion":        "Apropiación Social (ASC) > 0 o Divulgación (DPC) > 0",
            "cumple":           cumple_asc_dpc,
            "actual":           round(asc + dpc, 3),
            "objetivo_q1":      0.001,
            "objetivo_min_a1":  stats_a1.get("ASC", {}).get("min", 0)
                                or stats_a1.get("DPC", {}).get("min", 0),
            "brecha":           0 if cumple_asc_dpc else 1,
            "unidades_necesarias": (
                "" if cumple_asc_dpc else
                "Al menos 1 producto de apropiación social o divulgación"
            ),
        })

        # ── 4. FRH_A > 0 ──────────────────────────────────────────────────
        frh_a = sub_indicadores.get("FRH_A", 0)
        cumple_frh = frh_a > 0
        brechas.append({
            "condicion":        "Formación RH Tipo A (FRH_A) > 0",
            "cumple":           cumple_frh,
            "actual":           round(frh_a, 3),
            "objetivo_q1":      0.001,
            "objetivo_min_a1":  stats_a1.get("FRH_A", {}).get("min", 0),
            "brecha":           0 if cumple_frh else 1,
            "unidades_necesarias": (
                "" if cumple_frh else
                "Al menos 1 tesis doctoral o maestría dirigida"
            ),
        })

        # ── 5. Cohesión > 0 ────────────────────────────────────────────────
        coh = sub_indicadores.get("cohesion", 0)
        cumple_coh = coh > 0
        brechas.append({
            "condicion":        "Cohesión > 0",
            "cumple":           cumple_coh,
            "actual":           round(coh, 3),
            "objetivo_q1":      0.001,
            "objetivo_min_a1":  stats_a1.get("cohesion", {}).get("min", 0),
            "brecha":           0 if cumple_coh else 1,
            "unidades_necesarias": (
                "" if cumple_coh else
                "Publicaciones con co-autoría entre integrantes del grupo"
            ),
        })

        return brechas

    def _cuartiles_para_indicador(self, area: str, nombre_cuartil_contains: str) -> dict | None:
        grupos = set(self.grupos_en_area(area))
        df = self._df_cuartiles[
            self._df_cuartiles["grupo"].isin(grupos)
            & self._df_cuartiles["cuartil"].str.contains(
                nombre_cuartil_contains, case=False, na=False)
        ]
        if df.empty:
            return None
        return {
            "min": float(df["min"].median()),
            "q4":  float(df["q4"].median()),
            "q3":  float(df["q3"].median()),
            "q2":  float(df["q2"].median()),
            "max": float(df["max"].median()),
        }

    def _estimar_articulos_para_ig(self, delta_ig: float, maximos: dict) -> str:
        """Estima cuántos artículos Q1 se necesitarían para cerrar la brecha de IG."""
        max_top = maximos.get("NC_TOP", 0)
        if max_top <= 0:
            return ""
        # Cada artículo Q1 (lambda=1.0) agrega: (1.0/max_top) × 3.7 al IG
        tasa = (1.0 / max_top) * 3.7
        if tasa <= 0:
            return ""
        n = int(delta_ig / tasa) + 1
        return f"≈{n} artículos Q1 en TOP (o combinación equivalente)"


# Claves de sub-indicadores que van a df_indicadores
_CUARTIL_A_IND_KEYS = ["NC_TOP", "NC_A", "NC_B", "ASC", "DPC", "FRH_A", "FRH_B"]


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    pdf_dir = sys.argv[1] if len(sys.argv) > 1 else None
    out     = sys.argv[2] if len(sys.argv) > 2 else "auto"
    if out == "none":
        out = None

    dfs = parse_and_export(pdf_dir=pdf_dir, out=out)
    print("\n── Resumen ──────────────────────────────────────────────────")
    for nombre, df in dfs.items():
        print(f"  {nombre:15s}: {len(df):5d} filas, {len(df.columns)} columnas")
    print()
    if not dfs["grupos"].empty:
        print("Grupos por categoría:")
        print(dfs["grupos"].groupby("categoria").size().to_string())
