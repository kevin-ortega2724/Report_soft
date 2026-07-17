"""
vista_clasificacion.py

UI separada (fuera de la app principal, se integrará más adelante) para ver,
por grupo: la categoría VIGENTE (scraping GrupLAC, 100% técnico y actual) y,
para los grupos con línea base oficial, una PROYECCIÓN ESTIMADA de cómo
quedarían si la convocatoria se recalculara hoy con su producción de
artículos actual (ver proyeccion_957.py para el alcance exacto y sus
límites documentados).

Al seleccionar un grupo en la tabla, el panel de detalle muestra el cálculo
metodológico completo (indicadores, cuartiles del área, condiciones oficiales
de MinCiencias por categoría, y el detalle de qué artículos se contaron en
la proyección) -- no solo la categoría final, para que se pueda auditar de
dónde sale cada número, igual que el documento oficial (ver
docs/metodologia_957_plan_mejora.md y data_sim/Ficha-de-Asesoria...).

Módulo autocontenido: importa de src/ lo ya construido y validado
(Simulador957, build_categorias_grupos_957, DatabaseManager) en vez de
reimplementarlo.
"""

from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path
from typing import Optional

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))
sys.path.insert(0, str(Path(__file__).parent))

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QAbstractItemView, QHBoxLayout, QHeaderView, QLabel, QPushButton,
    QSplitter, QTableWidget, QTableWidgetItem, QTabWidget, QTextEdit,
    QVBoxLayout, QWidget,
)

from constants import CUARTIL_OBJETIVO_POR_CATEGORIA
from build_categorias_grupos_957 import construir_cache
from simulador_957 import Simulador957, ORDEN_CATEGORIAS
from proyeccion_957 import ProyeccionPublindex, proyectar_todos

_ORDEN_CATEGORIA = {"A1": 0, "A": 1, "B": 2, "C": 3, "Reconocido": 4, "Sin clasificar": 5}
_COLOR_CATEGORIA = {
    "A1": "#0ca30c",
    "A": "#4a90d9",
    "B": "#fab219",
    "C": "#ec835a",
    "Reconocido": "#9b9b9b",
    "Sin clasificar": "#d0d0d0",
}

_ETIQUETA_INDICADOR = {
    "NC_TOP": "Nuevo Conocimiento TOP",
    "NC_A": "Nuevo Conocimiento A",
    "NC_B": "Nuevo Conocimiento B",
    "ASC": "Apropiación Social del Conocimiento",
    "DPC": "Divulgación Pública de la Ciencia",
    "FRH_A": "Formación de Recurso Humano A",
    "FRH_B": "Formación de Recurso Humano B",
    "cohesion": "Cohesión",
    "colaboracion": "Colaboración",
}
_ORDEN_INDICADORES = ["NC_TOP", "NC_A", "NC_B", "ASC", "DPC", "FRH_A", "FRH_B", "cohesion", "colaboracion"]


def _fmt(v, decimales=2):
    if v is None:
        return "—"
    try:
        return f"{v:,.{decimales}f}"
    except (TypeError, ValueError):
        return str(v)


def _check(cumple: bool) -> str:
    color = "#0ca30c" if cumple else "#d03b3b"
    simbolo = "✓" if cumple else "✗"
    return f"<span style='color:{color};font-weight:bold;'>{simbolo}</span>"


class VistaClasificacion957(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self._canvas_actual = None
        self._sim: Optional[Simulador957] = None
        self._proyeccion_por_bd: dict[str, dict] = {}
        self._filas_por_grupo: dict[str, dict] = {}
        self.setup_ui()
        self.actualizar()

    # ── UI ───────────────────────────────────────────────────────────────
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        header = QHBoxLayout()
        titulo = QLabel("<b>Clasificación 957 — vigente y proyección estimada</b>")
        titulo.setStyleSheet("color:#1a365d; font-size:13px;")
        header.addWidget(titulo)
        header.addStretch()
        self.btn_actualizar = QPushButton("Actualizar")
        self.btn_actualizar.setToolTip(
            "Vuelve a leer la categoría vigente del scrape GrupLAC más reciente "
            "y a correr la proyección estimada.")
        self.btn_actualizar.clicked.connect(self.actualizar)
        header.addWidget(self.btn_actualizar)
        layout.addLayout(header)

        nota = QLabel(
            "\"Categoría vigente\": la reportada por GrupLAC/MinCiencias en el perfil de cada "
            "grupo, válida hasta la próxima convocatoria (100% técnico, siempre actualizado). "
            "\"Categoría proyectada\": ESTIMACIÓN (no oficial) de cómo quedaría el grupo si la "
            "convocatoria se recalculara hoy, usando el motor validado Simulador957 alimentado "
            "con los artículos publicados después del cierre de la ventana oficial (2023), "
            "clasificados por ISSN contra el dataset abierto de Publindex. Solo se proyectan "
            "artículos (NC_TOP/NC_A/NC_B) -- libros, extensión, divulgación y formación de "
            "recurso humano quedan en su valor oficial. Solo disponible para los grupos con "
            "línea base oficial (medicion_957.xlsx). Selecciona un grupo en la tabla para ver "
            "el cálculo completo (indicadores, cuartiles y condiciones oficiales de MinCiencias)."
        )
        nota.setWordWrap(True)
        nota.setStyleSheet(
            "color:#8a6d00; background:#fff8e1; border:1px solid #f0d98c; "
            "border-radius:4px; padding:6px 8px; font-size:10px;")
        layout.addWidget(nota)

        self.lbl_estado = QLabel("")
        self.lbl_estado.setStyleSheet("color:#898781; font-size:10px;")
        layout.addWidget(self.lbl_estado)

        fila_stats = QHBoxLayout()
        self._chips: dict[str, QLabel] = {}
        for clave in ("Total", "A1", "A", "B", "C", "Reconocido", "Sin clasificar", "Con proyección"):
            chip = QLabel(f"{clave}: —")
            chip.setStyleSheet(
                "background:#eaf1f8; color:#1a365d; border:1px solid #c7d9ea; "
                "border-radius:8px; padding:8px 12px; font-weight:bold; font-size:11px;")
            self._chips[clave] = chip
            fila_stats.addWidget(chip)
        fila_stats.addStretch()
        layout.addLayout(fila_stats)

        splitter = QSplitter(Qt.Horizontal)

        self.tabla = QTableWidget()
        self.tabla.setColumnCount(8)
        self.tabla.setHorizontalHeaderLabels([
            "Grupo", "Categoría vigente", "Categoría oficial (base)", "Categoría proyectada",
            "Artículos nuevos considerados", "Área de conocimiento", "Facultad", "Integrantes",
        ])
        self.tabla.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tabla.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.tabla.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tabla.setAlternatingRowColors(True)
        self.tabla.setSortingEnabled(True)
        self.tabla.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tabla.itemSelectionChanged.connect(self._mostrar_detalle_seleccion)
        splitter.addWidget(self.tabla)

        self.tabs_derecha = QTabWidget()

        self.panel_grafico = QWidget()
        self._lay_grafico = QVBoxLayout(self.panel_grafico)
        self.tabs_derecha.addTab(self.panel_grafico, "Distribución general")

        self.texto_detalle = QTextEdit()
        self.texto_detalle.setReadOnly(True)
        self.texto_detalle.setStyleSheet("background:#ffffff; font-size:11px;")
        self.texto_detalle.setHtml(
            "<p style='color:#898781;'>Selecciona un grupo en la tabla para ver el cálculo "
            "metodológico completo.</p>")
        self.tabs_derecha.addTab(self.texto_detalle, "Detalle del grupo")

        splitter.addWidget(self.tabs_derecha)
        splitter.setSizes([850, 550])

        layout.addWidget(splitter, 1)

    # ── Datos ────────────────────────────────────────────────────────────
    def actualizar(self):
        self.lbl_estado.setText("Cargando (leyendo scrape GrupLAC + corriendo proyección)...")
        self.repaint()

        categorias_gruplac = construir_cache()

        cursor = self.db.conn.cursor()
        tamanos = dict(cursor.execute('''
            SELECT grupo, COUNT(DISTINCT cedula) FROM grupos
            WHERE grupo IS NOT NULL AND grupo != '' GROUP BY grupo
        ''').fetchall())

        conteo_facultad: dict[str, tuple[str, int]] = {}
        for grupo, facultad, cnt in cursor.execute('''
            SELECT grupo, facultad, COUNT(*) FROM grupos
            WHERE grupo IS NOT NULL AND grupo != '' AND facultad IS NOT NULL AND facultad != ''
            GROUP BY grupo, facultad
        ''').fetchall():
            actual = conteo_facultad.get(grupo)
            if actual is None or cnt > actual[1]:
                conteo_facultad[grupo] = (facultad, cnt)
        facultades = {g: v[0] for g, v in conteo_facultad.items()}

        self._proyeccion_por_bd = {}
        try:
            self._sim = Simulador957()
            proyector = ProyeccionPublindex(self.db)
            proyeccion = proyectar_todos(self._sim, proyector)
            for nombre_medicion, datos in proyeccion.items():
                nombre_bd = proyector._nombre_en_bd(nombre_medicion)
                if nombre_bd:
                    self._proyeccion_por_bd[nombre_bd] = datos
            estado_proyeccion = f"{len(proyeccion)} grupos con línea base oficial evaluados."
        except Exception as e:
            estado_proyeccion = f"Proyección no disponible ({e})."

        filas = []
        conteo_cat: Counter = Counter()
        for grupo, info in categorias_gruplac.items():
            cat_vigente = info.get('categoria_asignada') or 'Sin clasificar'
            conteo_cat[cat_vigente] += 1
            proy = self._proyeccion_por_bd.get(grupo)
            filas.append({
                'grupo': grupo,
                'categoria_vigente': cat_vigente,
                'categoria_oficial_base': proy['categoria_oficial'] if proy else None,
                'categoria_proyectada': proy['categoria_proyectada'] if proy else None,
                'articulos_nuevos': proy['n_articulos_nuevos_clasificados'] if proy else None,
                'area': info.get('area_conocimiento') or '',
                'facultad': facultades.get(grupo, ''),
                'integrantes': tamanos.get(grupo, 0),
            })

        self._filas_por_grupo = {f['grupo']: f for f in filas}
        self._poblar_tabla(filas)
        self._actualizar_chips(conteo_cat, len(filas), len(self._proyeccion_por_bd))
        self._graficar_distribucion(conteo_cat)
        self.lbl_estado.setText(
            f"Categoría vigente: scrape GrupLAC más reciente. Proyección: {estado_proyeccion}")

    # ── UI: tabla / chips / gráfica ──────────────────────────────────────
    def _poblar_tabla(self, filas):
        filas = sorted(filas, key=lambda f: (
            _ORDEN_CATEGORIA.get(f['categoria_vigente'], 9), f['grupo']))

        self.tabla.setSortingEnabled(False)
        self.tabla.setRowCount(len(filas))
        for row, f in enumerate(filas):
            item_grupo = QTableWidgetItem(f['grupo'])

            item_vigente = QTableWidgetItem(f['categoria_vigente'])
            item_vigente.setBackground(QColor(_COLOR_CATEGORIA.get(f['categoria_vigente'], "#ffffff")))

            item_oficial = QTableWidgetItem(f['categoria_oficial_base'] or "—")

            proyectada = f['categoria_proyectada']
            item_proyectada = QTableWidgetItem(proyectada or "—")
            if proyectada and f['categoria_oficial_base'] and proyectada != f['categoria_oficial_base']:
                item_proyectada.setBackground(QColor("#fff3cd"))  # resalta cambio proyectado

            n_nuevos = f['articulos_nuevos']
            item_nuevos = QTableWidgetItem("—" if n_nuevos is None else str(n_nuevos))
            if n_nuevos is not None:
                item_nuevos.setData(Qt.DisplayRole, n_nuevos)

            item_area = QTableWidgetItem(f['area'])
            item_fac = QTableWidgetItem(f['facultad'])
            item_int = QTableWidgetItem(str(f['integrantes']))
            item_int.setData(Qt.DisplayRole, f['integrantes'])

            for col, item in enumerate((
                    item_grupo, item_vigente, item_oficial, item_proyectada,
                    item_nuevos, item_area, item_fac, item_int)):
                self.tabla.setItem(row, col, item)
        self.tabla.setSortingEnabled(True)

    def _actualizar_chips(self, conteo_cat, total, con_proyeccion):
        self._chips["Total"].setText(f"Total: {total}")
        for clave in ("A1", "A", "B", "C", "Reconocido", "Sin clasificar"):
            self._chips[clave].setText(f"{clave}: {conteo_cat.get(clave, 0)}")
        self._chips["Con proyección"].setText(f"Con proyección: {con_proyeccion}")

    def _graficar_distribucion(self, conteo_cat):
        from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
        from matplotlib.figure import Figure

        if self._canvas_actual is not None:
            self._lay_grafico.removeWidget(self._canvas_actual)
            self._canvas_actual.setParent(None)

        etiquetas = [c for c in _ORDEN_CATEGORIA if conteo_cat.get(c, 0) > 0]
        valores = [conteo_cat[c] for c in etiquetas]
        colores = [_COLOR_CATEGORIA[c] for c in etiquetas]

        fig = Figure(figsize=(4.5, 4), facecolor="#fcfcfb")
        ax = fig.add_subplot(111)
        ax.set_facecolor("#fcfcfb")
        x = list(range(len(etiquetas)))
        ax.bar(x, valores, color=colores, zorder=3)
        ax.set_xticks(x)
        ax.set_xticklabels(etiquetas, fontsize=9)
        for i, v in enumerate(valores):
            ax.text(i, v + 0.3, str(v), ha="center", fontsize=9, color="#52514e")
        ax.set_title("Distribución de categorías vigentes", fontsize=9, color="#1a365d")
        ax.grid(axis="y", color="#e1e0d9", linewidth=1, zorder=0)
        ax.set_axisbelow(True)
        for spine in ("top", "right"):
            ax.spines[spine].set_visible(False)
        fig.tight_layout()

        self._canvas_actual = FigureCanvas(fig)
        self._lay_grafico.addWidget(self._canvas_actual)

    # ── Detalle metodológico por grupo ───────────────────────────────────
    def _mostrar_detalle_seleccion(self):
        seleccion = self.tabla.selectedItems()
        if not seleccion:
            return
        fila = seleccion[0].row()
        grupo = self.tabla.item(fila, 0).text()
        self.texto_detalle.setHtml(self._html_detalle(grupo))
        self.tabs_derecha.setCurrentWidget(self.texto_detalle)

    def _html_detalle(self, grupo: str) -> str:
        info = self._filas_por_grupo.get(grupo, {})
        proy = self._proyeccion_por_bd.get(grupo)

        html = [
            f"<h3 style='color:#1a365d;margin-bottom:2px;'>{grupo}</h3>",
            f"<p style='color:#555;font-size:10px;margin-top:0;'>"
            f"Área: {info.get('area') or '—'} · Facultad: {info.get('facultad') or '—'} · "
            f"Integrantes: {info.get('integrantes', '—')}</p>",
            f"<p><b>Categoría vigente (GrupLAC):</b> "
            f"<span style='background:{_COLOR_CATEGORIA.get(info.get('categoria_vigente'), '#ddd')};"
            f"color:white;padding:2px 8px;border-radius:3px;'>{info.get('categoria_vigente')}</span></p>",
        ]

        if proy is None:
            html.append(
                "<div style='background:#f0f0f0;border-radius:4px;padding:8px;color:#666;'>"
                "Este grupo no tiene línea base oficial en <code>medicion_957.xlsx</code> "
                "(medición Convocatoria 957) -- sin ese dato no hay <code>ratio</code> que "
                "calibrar (ver <code>simulador_957.py</code>), así que no se puede proyectar "
                "ni mostrar el desglose de indicadores/cuartiles. Solo se muestra la categoría "
                "vigente reportada por GrupLAC.</div>"
            )
            return "".join(html)

        oficial = proy["oficial"]
        proyectado = proy["proyectado"]
        area = oficial.area
        cuartiles_area = (self._sim._cuartiles_area.get(area, {}) if self._sim else {})

        html.append(
            f"<p><b>Categoría oficial (medición 957):</b> {oficial.categoria_oficial or '—'} "
            f"&nbsp;→&nbsp; <b>Categoría proyectada (estimada):</b> {proyectado.categoria_simulada}"
            + (" <span style='color:#8a6d00;'>(cambio estimado)</span>"
               if oficial.categoria_simulada != proyectado.categoria_simulada else "")
            + "</p>"
        )

        # ── Paso 5-6: IG y cuartiles ─────────────────────────────────────
        html.append("<h4 style='color:#1a365d;margin-bottom:2px;'>Índice Global (Paso 5) y cuartil (Paso 6)</h4>")
        html.append(
            f"<p style='font-size:10px;color:#555;margin-top:0;'>IG = Σ (ponderación × índice) por "
            f"indicador, índice = valor_indicador / máximo del área.</p>"
        )
        html.append(
            "<table style='border-collapse:collapse;font-size:10px;' width='100%'>"
            "<tr style='background:#1a365d;color:white;'>"
            "<th style='padding:4px;'>&nbsp;</th><th style='padding:4px;'>Oficial</th>"
            "<th style='padding:4px;'>Proyectado</th></tr>"
            f"<tr><td style='padding:4px;'>IG (Índice Global)</td>"
            f"<td style='padding:4px;text-align:center;'>{_fmt(oficial.ig_oficial)}</td>"
            f"<td style='padding:4px;text-align:center;'>{_fmt(proyectado.ig_simulado)}</td></tr>"
            f"<tr><td style='padding:4px;'>Cuartil del IG (1=25% superior)</td>"
            f"<td style='padding:4px;text-align:center;'>{oficial.cuartil_ig}</td>"
            f"<td style='padding:4px;text-align:center;'>{proyectado.cuartil_ig}</td></tr>"
            f"<tr><td style='padding:4px;'>Cuartil de NC_TOP</td>"
            f"<td style='padding:4px;text-align:center;'>{oficial.cuartil_nc_top}</td>"
            f"<td style='padding:4px;text-align:center;'>{proyectado.cuartil_nc_top}</td></tr>"
            f"<tr><td style='padding:4px;'>Años de existencia (a la fecha de la convocatoria)</td>"
            f"<td colspan='2' style='padding:4px;text-align:center;'>{_fmt(oficial.anios_existencia, 0)}</td></tr>"
            "</table>"
        )

        # ── Indicadores + cuartiles del área ─────────────────────────────
        html.append(
            "<h4 style='color:#1a365d;margin:8px 0 2px;'>Indicadores vs. distribución nacional del área</h4>")
        html.append(
            f"<p style='font-size:10px;color:#555;margin-top:0;'>Umbrales min/q4/q3/q2/max de "
            f"\"{area}\" (hoja <code>cuartiles</code> de medicion_957.xlsx, 100% oficial): "
            f"q4≈percentil 25, q3≈percentil 50, q2≈percentil 75 de los grupos del área.</p>"
        )
        html.append(
            "<table style='border-collapse:collapse;font-size:9px;' width='100%'>"
            "<tr style='background:#1a365d;color:white;'>"
            "<th style='padding:3px;'>Indicador</th><th style='padding:3px;'>Oficial</th>"
            "<th style='padding:3px;'>Proyectado</th><th style='padding:3px;'>min</th>"
            "<th style='padding:3px;'>q4</th><th style='padding:3px;'>q3</th>"
            "<th style='padding:3px;'>q2</th><th style='padding:3px;'>max</th></tr>"
        )
        for ind in _ORDEN_INDICADORES:
            v_of = oficial.valores_simulados.get(ind)
            v_pr = proyectado.valores_simulados.get(ind)
            umbrales = cuartiles_area.get(ind, {})
            resaltar = "font-weight:bold;" if v_pr is not None and v_of is not None and v_pr != v_of else ""
            html.append(
                f"<tr><td style='padding:3px;'>{_ETIQUETA_INDICADOR.get(ind, ind)}</td>"
                f"<td style='padding:3px;text-align:center;'>{_fmt(v_of)}</td>"
                f"<td style='padding:3px;text-align:center;{resaltar}'>{_fmt(v_pr)}</td>"
                f"<td style='padding:3px;text-align:center;'>{_fmt(umbrales.get('min'))}</td>"
                f"<td style='padding:3px;text-align:center;'>{_fmt(umbrales.get('q4'))}</td>"
                f"<td style='padding:3px;text-align:center;'>{_fmt(umbrales.get('q3'))}</td>"
                f"<td style='padding:3px;text-align:center;'>{_fmt(umbrales.get('q2'))}</td>"
                f"<td style='padding:3px;text-align:center;'>{_fmt(umbrales.get('max'))}</td></tr>"
            )
        html.append("</table>")

        # ── Condiciones oficiales por categoría (Paso 6.2) ───────────────
        html.append(
            "<h4 style='color:#1a365d;margin:8px 0 2px;'>Condiciones oficiales por categoría (Paso 6.2)</h4>")
        html.append(
            "<p style='font-size:10px;color:#555;margin-top:0;'>Transcritas literalmente de los "
            "documentos oficiales de MinCiencias (data/pdf/A1s.pdf, As.pdf, Bs.pdf, Cs.pdf). "
            "Una categoría se alcanza cuando TODAS sus condiciones se cumplen. La condición de "
            "\"investigador emérito/sénior/asociado vinculado\" no se evalúa -- ese dato no existe "
            "en el scraping de GrupLAC a nivel de grupo (requiere el CvLAC individual).</p>"
        )
        for cat in reversed(ORDEN_CATEGORIAS[1:]):  # A1, A, B, C
            cond_of = oficial.condiciones.get(cat, {})
            cond_pr = proyectado.condiciones.get(cat, {})
            if not cond_of:
                continue
            cumple_of_todas = all(cond_of.values())
            etiqueta_alcanzada = (
                "<span style='color:#0ca30c;'>— alcanzada (oficial)</span>" if cumple_of_todas else ""
            )
            html.append(
                f"<p style='margin:6px 0 2px;'><b>Categoría {cat}</b> {etiqueta_alcanzada}</p>"
            )
            html.append(
                "<table style='border-collapse:collapse;font-size:9px;margin-bottom:4px;' width='100%'>"
                "<tr style='background:#eef2f6;'><th style='padding:2px;text-align:left;'>Condición</th>"
                "<th style='padding:2px;width:60px;'>Oficial</th><th style='padding:2px;width:70px;'>Proyectada</th></tr>"
            )
            for condicion, cumple in cond_of.items():
                cumple_pr = cond_pr.get(condicion, cumple)
                html.append(
                    f"<tr><td style='padding:2px;'>{condicion}</td>"
                    f"<td style='padding:2px;text-align:center;'>{_check(cumple)}</td>"
                    f"<td style='padding:2px;text-align:center;'>{_check(cumple_pr)}</td></tr>"
                )
            html.append("</table>")

        # ── Detalle de la proyección: artículos considerados ─────────────
        detalle_articulos = proy.get("detalle_articulos") or []
        html.append(
            "<h4 style='color:#1a365d;margin:8px 0 2px;'>Artículos nuevos considerados en la proyección</h4>")
        if not detalle_articulos:
            html.append(
                "<p style='font-size:10px;color:#666;'>Ninguno -- no se encontraron publicaciones "
                "internas con ISSN y año posterior a 2023 para este grupo.</p>")
        else:
            html.append(
                "<p style='font-size:10px;color:#555;margin-top:0;'>Publicados después del cierre de "
                "la ventana oficial (2023), clasificados por ISSN contra Publindex "
                f"(tope de vigencia disponible: 2022). {len(detalle_articulos)} encontrados, "
                f"{sum(1 for a in detalle_articulos if a['subtipo'])} con categoría Publindex resuelta.</p>"
            )
            html.append(
                "<table style='border-collapse:collapse;font-size:9px;' width='100%'>"
                "<tr style='background:#1a365d;color:white;'>"
                "<th style='padding:3px;'>Título</th><th style='padding:3px;'>ISSN</th>"
                "<th style='padding:3px;'>Año</th><th style='padding:3px;'>Categoría Publindex</th>"
                "<th style='padding:3px;'>¿Contado?</th></tr>"
            )
            for a in detalle_articulos:
                cat_pub = a["categoria_publindex"] or "no encontrado en Publindex"
                color_fila = "" if a["subtipo"] else "color:#999;font-style:italic;"
                html.append(
                    f"<tr style='{color_fila}'><td style='padding:3px;'>{(a['titulo'] or '')[:70]}</td>"
                    f"<td style='padding:3px;text-align:center;'>{a['issn']}</td>"
                    f"<td style='padding:3px;text-align:center;'>{a['anio']}</td>"
                    f"<td style='padding:3px;text-align:center;'>{cat_pub}</td>"
                    f"<td style='padding:3px;text-align:center;'>{_check(bool(a['subtipo']))}</td></tr>"
                )
            html.append("</table>")

            if proyectado.detalle_ajustes:
                html.append(
                    "<p style='font-size:10px;color:#555;margin:6px 0 2px;'><b>λ (lambda) antes/después "
                    "por subtipo</b> -- λ = ln(1 + total / ventana):</p>")
                html.append(
                    "<table style='border-collapse:collapse;font-size:9px;' width='100%'>"
                    "<tr style='background:#eef2f6;'><th style='padding:2px;'>Subtipo</th>"
                    "<th style='padding:2px;'>Ventana</th><th style='padding:2px;'>Total antes</th>"
                    "<th style='padding:2px;'>Total simulado</th><th style='padding:2px;'>λ antes</th>"
                    "<th style='padding:2px;'>λ simulado</th><th style='padding:2px;'>Δλ</th></tr>"
                )
                for subtipo, d in proyectado.detalle_ajustes.items():
                    html.append(
                        f"<tr><td style='padding:2px;'>{subtipo}</td>"
                        f"<td style='padding:2px;text-align:center;'>{d['ventana']}</td>"
                        f"<td style='padding:2px;text-align:center;'>{d['total_actual']}</td>"
                        f"<td style='padding:2px;text-align:center;'>{d['total_simulado']}</td>"
                        f"<td style='padding:2px;text-align:center;'>{d['lambda_actual']}</td>"
                        f"<td style='padding:2px;text-align:center;'>{d['lambda_simulado']}</td>"
                        f"<td style='padding:2px;text-align:center;'>{d['delta_lambda']}</td></tr>"
                    )
                html.append("</table>")

        return "".join(html)
