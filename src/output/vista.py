#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
visualizar_gruplac.py
Lee:
- SQLite: gruplac_clasificacion.db
- TXT: reporte_clasificacion.txt
Y genera un dashboard HTML interactivo con Plotly.

Uso:
  python visualizar_gruplac.py --db gruplac_clasificacion.db --report reporte_clasificacion.txt --outdir salida
  python visualizar_gruplac.py --group "GESTIÓN DE SISTEMAS ELÉCTRICOS, ELECTRÓNICOS Y AUTOMÁTICOS"
"""

from __future__ import annotations

import argparse
import os
import re
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go


# ----------------------------
# Utilidades
# ----------------------------

def safe_mkdir(path: str) -> None:
    os.makedirs(path, exist_ok=True)

def normalize_ws(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def to_int(x, default=0) -> int:
    try:
        return int(x)
    except Exception:
        return default

def now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


# ----------------------------
# Lectura y parsing del TXT
# ----------------------------

@dataclass
class ReportSummary:
    fecha: Optional[str]
    total_grupos: Optional[int]
    total_registros: Optional[int]
    global_por_categoria: Dict[str, int]
    pestanas_no_clasificadas_global: List[str]


def parse_report_txt(report_path: str) -> ReportSummary:
    """
    Parser robusto para reporte_clasificacion.txt.
    Extrae:
      - Fecha
      - Total de grupos procesados
      - Total de registros de productos
      - Resumen global por categoría (5 categorías)
      - Lista de pestañas no clasificadas (global)
    """
    if not os.path.exists(report_path):
        return ReportSummary(
            fecha=None,
            total_grupos=None,
            total_registros=None,
            global_por_categoria={},
            pestanas_no_clasificadas_global=[],
        )

    text = open(report_path, "r", encoding="utf-8", errors="ignore").read()

    fecha = None
    m = re.search(r"Fecha:\s*([0-9:\-\s]+)", text)
    if m:
        fecha = normalize_ws(m.group(1))

    total_grupos = None
    m = re.search(r"Total de grupos procesados:\s*(\d+)", text)
    if m:
        total_grupos = to_int(m.group(1), None)

    total_registros = None
    m = re.search(r"Total de registros de productos:\s*(\d+)", text)
    if m:
        total_registros = to_int(m.group(1), None)

    # Resumen global por categoría (líneas tipo: "1. Generación ...: 13948 productos")
    global_cat = {}
    for line in text.splitlines():
        line_n = normalize_ws(line)
        m = re.match(r"^\d+\.\s*(.+?):\s*([0-9]+)\s*productos$", line_n, flags=re.IGNORECASE)
        if m:
            cat = normalize_ws(m.group(1))
            val = to_int(m.group(2), 0)
            global_cat[cat] = val

    # Pestañas no clasificadas (global) -> bloque después del título
    pestanas = []
    block = re.search(
        r"PESTAÑAS NO CLASIFICADAS.*?\n(-+\n)(.*?)(\n-+\n|\nDETALLE POR GRUPO)",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if block:
        lines = block.group(2).splitlines()
        for ln in lines:
            ln = normalize_ws(ln)
            if ln.startswith("-"):
                pestanas.append(normalize_ws(ln.lstrip("-").strip()))
    pestanas = [p for p in pestanas if p]

    return ReportSummary(
        fecha=fecha,
        total_grupos=total_grupos,
        total_registros=total_registros,
        global_por_categoria=global_cat,
        pestanas_no_clasificadas_global=pestanas,
    )


# ----------------------------
# Lectura de la BD SQLite
# ----------------------------

def read_sqlite_tables(db_path: str) -> Dict[str, pd.DataFrame]:
    if not os.path.exists(db_path):
        raise FileNotFoundError(f"No existe la BD: {db_path}")

    con = sqlite3.connect(db_path)
    try:
        tables = pd.read_sql_query(
            "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name",
            con
        )["name"].tolist()

        out: Dict[str, pd.DataFrame] = {}
        for t in tables:
            # evita sqlite_sequence si no aporta
            if t == "sqlite_sequence":
                continue
            out[t] = pd.read_sql_query(f"SELECT * FROM {t}", con)
        return out
    finally:
        con.close()


# ----------------------------
# Visualizaciones
# ----------------------------

def fig_global_categories_from_db(df_resumen: pd.DataFrame) -> go.Figure:
    """
    df_resumen: tabla resumen_categorias (grupo_id, grupo_nombre, categoria, subcategoria, total_productos)
    Agrega por categoria global.
    """
    d = (
        df_resumen
        .groupby("categoria", as_index=False)["total_productos"]
        .sum()
        .sort_values("total_productos", ascending=False)
    )
    fig = px.bar(d, x="categoria", y="total_productos", title="Resumen global por categoría (desde BD)")
    fig.update_layout(xaxis_title="Categoría Minciencias", yaxis_title="Total productos", xaxis_tickangle=-25)
    return fig

def fig_global_categories_from_report(rep: ReportSummary) -> go.Figure:
    d = pd.DataFrame(
        [{"categoria": k, "total_productos": v} for k, v in rep.global_por_categoria.items()]
    ).sort_values("total_productos", ascending=False)
    if d.empty:
        fig = go.Figure()
        fig.update_layout(title="Resumen global por categoría (desde TXT) - sin datos")
        return fig
    fig = px.bar(d, x="categoria", y="total_productos", title="Resumen global por categoría (desde TXT)")
    fig.update_layout(xaxis_title="Categoría (TXT)", yaxis_title="Total productos", xaxis_tickangle=-25)
    return fig

def fig_top_groups(df_grupos: pd.DataFrame, df_resumen: pd.DataFrame, top_n: int = 20) -> Tuple[go.Figure, go.Figure]:
    """
    Devuelve dos figuras:
      - Top grupos por total_productos (sum resumen_categorias)
      - Top grupos por registros totales (si existe en grupos o aproximación)
    """
    # Top por total_productos (desde resumen_categorias)
    dprod = (
        df_resumen
        .groupby("grupo_nombre", as_index=False)["total_productos"]
        .sum()
        .sort_values("total_productos", ascending=False)
        .head(top_n)
    )
    fig_prod = px.bar(dprod, x="total_productos", y="grupo_nombre", orientation="h",
                      title=f"Top {top_n} grupos por total de productos (BD)")
    fig_prod.update_layout(xaxis_title="Total productos", yaxis_title="Grupo", yaxis={"categoryorder": "total ascending"})

    # Registros totales: usa grupos.total_pestanas / pestanas_clasificadas (no es registros)
    # Mejor: aproximación desde productos.cantidad_registros si está.
    # Si no, se deja el gráfico vacío.
    fig_reg = go.Figure()
    fig_reg.update_layout(title=f"Top {top_n} grupos por total de registros (BD)")

    return fig_prod, fig_reg

def fig_group_breakdown(
    grupo: str,
    df_grupos: pd.DataFrame,
    df_resumen: pd.DataFrame,
    df_productos: pd.DataFrame,
    df_pestanas: pd.DataFrame,
) -> Dict[str, go.Figure]:
    """
    Construye un set de figuras para un grupo específico.
    """
    grupo_norm = normalize_ws(grupo)

    # 1) Barras por categoría
    dg = df_resumen[df_resumen["grupo_nombre"].astype(str).str.strip() == grupo_norm].copy()
    if dg.empty:
        # intentamos match insensible
        mask = df_resumen["grupo_nombre"].astype(str).str.lower().str.strip() == grupo_norm.lower()
        dg = df_resumen[mask].copy()

    figs: Dict[str, go.Figure] = {}

    if dg.empty:
        fig = go.Figure()
        fig.update_layout(title=f"Grupo: {grupo_norm} (no encontrado en resumen_categorias)")
        figs["grupo_categoria"] = fig
        return figs

    dc = (
        dg.groupby("categoria", as_index=False)["total_productos"]
        .sum()
        .sort_values("total_productos", ascending=False)
    )
    fig_cat = px.bar(dc, x="categoria", y="total_productos",
                     title=f"Grupo: {grupo_norm} | Productos por categoría")
    fig_cat.update_layout(xaxis_title="Categoría", yaxis_title="Total productos", xaxis_tickangle=-25)
    figs["grupo_categoria"] = fig_cat

    # 2) Top subcategorías del grupo
    ds = (
        dg.groupby(["categoria", "subcategoria"], as_index=False)["total_productos"]
        .sum()
        .sort_values("total_productos", ascending=False)
        .head(30)
    )
    fig_sub = px.bar(ds, x="total_productos", y="subcategoria", color="categoria", orientation="h",
                     title=f"Grupo: {grupo_norm} | Top 30 subcategorías (por productos)")
    fig_sub.update_layout(xaxis_title="Total productos", yaxis_title="Subcategoría", yaxis={"categoryorder": "total ascending"})
    figs["grupo_subcategorias"] = fig_sub

    # 3) Pestañas no clasificadas (desde tabla pestanas_disponibles)
    dp = df_pestanas[df_pestanas["grupo_nombre"].astype(str).str.strip() == grupo_norm].copy()
    if dp.empty:
        mask = df_pestanas["grupo_nombre"].astype(str).str.lower().str.strip() == grupo_norm.lower()
        dp = df_pestanas[mask].copy()

    if not dp.empty:
        dnc = dp[dp["clasificada"].astype(int) == 0].copy()
        if dnc.empty:
            fig_nc = go.Figure()
            fig_nc.update_layout(title=f"Grupo: {grupo_norm} | Pestañas no clasificadas (ninguna)")
        else:
            dnc["pestana"] = dnc["pestana"].astype(str)
            dnc_count = dnc.groupby("pestana", as_index=False).size().rename(columns={"size": "conteo"})
            fig_nc = px.bar(dnc_count.sort_values("conteo", ascending=False),
                            x="conteo", y="pestana", orientation="h",
                            title=f"Grupo: {grupo_norm} | Pestañas NO clasificadas")
            fig_nc.update_layout(xaxis_title="Conteo", yaxis_title="Pestaña", yaxis={"categoryorder": "total ascending"})
    else:
        fig_nc = go.Figure()
        fig_nc.update_layout(title=f"Grupo: {grupo_norm} | Pestañas (sin datos en pestanas_disponibles)")
    figs["grupo_pestanas_no_clas"] = fig_nc

    # 4) Productos reconocidos (tabla productos): top pestana_original o subcategoria
    dprod = df_productos[df_productos["grupo_nombre"].astype(str).str.strip() == grupo_norm].copy()
    if dprod.empty:
        mask = df_productos["grupo_nombre"].astype(str).str.lower().str.strip() == grupo_norm.lower()
        dprod = df_productos[mask].copy()

    if not dprod.empty:
        dprod2 = (
            dprod.groupby(["categoria_minciencias", "subcategoria"], as_index=False)["cantidad_registros"]
            .sum()
            .sort_values("cantidad_registros", ascending=False)
            .head(30)
        )
        fig_pr = px.bar(dprod2, x="cantidad_registros", y="subcategoria", color="categoria_minciencias",
                        orientation="h",
                        title=f"Grupo: {grupo_norm} | Top 30 subcategorías por #registros (tabla productos)")
        fig_pr.update_layout(xaxis_title="# Registros", yaxis_title="Subcategoría", yaxis={"categoryorder": "total ascending"})
    else:
        fig_pr = go.Figure()
        fig_pr.update_layout(title=f"Grupo: {grupo_norm} | Productos (sin datos en tabla productos)")
    figs["grupo_productos_registros"] = fig_pr

    return figs


# ----------------------------
# Dashboard HTML
# ----------------------------

def build_html_dashboard(
    out_html: str,
    rep: ReportSummary,
    tables: Dict[str, pd.DataFrame],
    group: Optional[str],
    top_n: int,
) -> None:
    df_grupos = tables.get("grupos", pd.DataFrame())
    df_resumen = tables.get("resumen_categorias", pd.DataFrame())
    df_productos = tables.get("productos", pd.DataFrame())
    df_pestanas = tables.get("pestanas_disponibles", pd.DataFrame())

    parts: List[str] = []

    # Header
    header = f"""
    <html>
    <head>
      <meta charset="utf-8"/>
      <title>Dashboard Gruplac - Clasificación</title>
      <style>
        body {{ font-family: Arial, sans-serif; margin: 18px; }}
        .meta {{ background: #f6f6f6; padding: 12px; border-radius: 10px; }}
        .grid {{ display: grid; grid-template-columns: 1fr; gap: 18px; }}
        .row2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }}
        .card {{ border: 1px solid #eee; border-radius: 12px; padding: 10px; }}
        h1, h2 {{ margin: 0.2em 0; }}
        ul {{ margin: 0.4em 0 0.2em 1.2em; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 6px; font-size: 12px; }}
        th {{ background: #fafafa; }}
      </style>
      <script src="https://cdn.plot.ly/plotly-2.30.0.min.js"></script>
    </head>
    <body>
      <h1>Reporte interactivo — Clasificación Gruplac (Convocatoria 957)</h1>
      <div class="meta">
        <b>Fecha (TXT):</b> {rep.fecha or "N/D"} &nbsp; | &nbsp;
        <b>Total grupos (TXT):</b> {rep.total_grupos if rep.total_grupos is not None else "N/D"} &nbsp; | &nbsp;
        <b>Total registros (TXT):</b> {rep.total_registros if rep.total_registros is not None else "N/D"}
        <br/>
        <b>Tablas en BD:</b> {", ".join(sorted(tables.keys()))}
      </div>
      <div class="grid">
    """
    parts.append(header)

    # Charts globales
    fig_db = fig_global_categories_from_db(df_resumen) if not df_resumen.empty else go.Figure()
    fig_txt = fig_global_categories_from_report(rep)

    parts.append('<div class="row2">')
    parts.append('<div class="card">' + fig_db.to_html(full_html=False, include_plotlyjs=False) + '</div>')
    parts.append('<div class="card">' + fig_txt.to_html(full_html=False, include_plotlyjs=False) + '</div>')
    parts.append('</div>')

    # Top grupos
    fig_top_prod, fig_top_reg = fig_top_groups(df_grupos, df_resumen, top_n=top_n)

    parts.append('<div class="row2">')
    parts.append('<div class="card">' + fig_top_prod.to_html(full_html=False, include_plotlyjs=False) + '</div>')
    parts.append('<div class="card">' + fig_top_reg.to_html(full_html=False, include_plotlyjs=False) + '</div>')
    parts.append('</div>')

    # Tabla resumen por grupo (top_n)
    if not df_resumen.empty:
        grp = (
            df_resumen.groupby(["grupo_nombre"], as_index=False)["total_productos"]
            .sum()
            .sort_values("total_productos", ascending=False)
            .head(top_n)
        )
        parts.append('<div class="card">')
        parts.append(f"<h2>Top {top_n} grupos — tabla (productos totales)</h2>")
        parts.append(grp.to_html(index=False, escape=True))
        parts.append("</div>")

    # Bloque de pestañas no clasificadas global (desde TXT)
    parts.append('<div class="card">')
    parts.append("<h2>Pestañas NO clasificadas (global, desde TXT)</h2>")
    if rep.pestanas_no_clasificadas_global:
        parts.append("<ul>")
        for p in rep.pestanas_no_clasificadas_global:
            parts.append(f"<li>{p}</li>")
        parts.append("</ul>")
    else:
        parts.append("<p>(Sin datos en TXT o no encontradas.)</p>")
    parts.append("</div>")

    # Detalle de un grupo (si se pide)
    if group:
        figs = fig_group_breakdown(group, df_grupos, df_resumen, df_productos, df_pestanas)
        parts.append('<div class="card">')
        parts.append(f"<h2>Detalle del grupo: {group}</h2>")
        parts.append("</div>")

        # Inserta figuras del grupo
        for k, fig in figs.items():
            parts.append('<div class="card">')
            parts.append(fig.to_html(full_html=False, include_plotlyjs=False))
            parts.append("</div>")

        # Además: snapshot de "grupos" para ese grupo
        if not df_grupos.empty:
            dg = df_grupos[df_grupos["nombre"].astype(str).str.strip().str.lower() == normalize_ws(group).lower()].copy()
            if not dg.empty:
                parts.append('<div class="card">')
                parts.append("<h2>Metadatos del grupo (tabla grupos)</h2>")
                parts.append(dg.to_html(index=False, escape=True))
                parts.append("</div>")

    # Footer
    parts.append("""
      </div>
      <hr/>
      <p style="font-size: 12px; color: #666;">
        Generado automáticamente por visualizar_gruplac.py
      </p>
    </body>
    </html>
    """)

    html = "\n".join(parts)
    with open(out_html, "w", encoding="utf-8") as f:
        f.write(html)


# ----------------------------
# Main
# ----------------------------

def main() -> int:
    ap = argparse.ArgumentParser(description="Visualizador Gruplac (BD + TXT) -> Dashboard HTML")
    ap.add_argument("--db", default="gruplac_clasificacion.db", help="Ruta a la BD SQLite")
    ap.add_argument("--report", default="reporte_clasificacion.txt", help="Ruta al TXT del reporte")
    ap.add_argument("--outdir", default="salida_gruplac_dashboard", help="Carpeta de salida")
    ap.add_argument("--group", default=None, help="Nombre exacto (o casi) del grupo para detalle")
    ap.add_argument("--top", type=int, default=20, help="Top N para tablas/gráficas")
    args = ap.parse_args()

    safe_mkdir(args.outdir)

    rep = parse_report_txt(args.report)
    tables = read_sqlite_tables(args.db)

    out_html = os.path.join(args.outdir, f"dashboard_gruplac_{now_stamp()}.html")
    build_html_dashboard(out_html, rep, tables, args.group, top_n=args.top)

    print("\n✅ Dashboard generado:")
    print(out_html)
    print("\nSugerencia: abre el archivo HTML en tu navegador.\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
