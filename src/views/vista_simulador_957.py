"""
Vista: Simulador oficial Conv. 957 (verificación + proyección).

Muestra, por grupo, el resultado oficial de MinCiencias (medicion_957.xlsx),
el detalle de productos y ventana de observación (Paso 2), los productos
crudos en GrupLAC para auditoría visual, y un simulador de impacto que
proyecta productos hipotéticos sobre los indicadores, el IG y la categoría.

Cubre solo los 75 grupos con PDF oficial de MinCiencias (los que tiene
`Simulador957`); no usa la BD interna de la aplicación.
"""

from collections import defaultdict

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtWidgets import (
    QComboBox, QGroupBox, QHBoxLayout, QHeaderView, QLabel, QMessageBox,
    QPushButton, QSpinBox, QTableWidget, QTableWidgetItem, QTabWidget,
    QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget,
)

from constants import COLORES_CATEGORIA
from simulador_957 import SECCIONES_CON_PRODUCTOS, Simulador957

ETIQUETAS_INDICADOR = {
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

COLOR_CATEGORIA_957 = {
    "A1": "#2E7D32",
    "A": "#558B2F",
    "B": "#F9A825",
    "C": "#E65100",
    "D": "#B71C1C",
}

BTN_PRIMARIO = (
    "QPushButton { background-color: #2E86AB; color: white; "
    "padding: 6px 16px; border-radius: 4px; font-weight: bold; }"
    "QPushButton:hover { background-color: #1a5276; }"
)
BTN_SECUNDARIO = (
    "QPushButton { background-color: #95a5a6; color: white; "
    "padding: 6px 16px; border-radius: 4px; }"
    "QPushButton:hover { background-color: #7f8c8d; }"
)


def _item(texto) -> QTableWidgetItem:
    it = QTableWidgetItem(str(texto))
    it.setFlags(it.flags() & ~Qt.ItemIsEditable)
    return it


# Color por procedencia del dato, para que cada pestaña deje explícito de
# dónde sale lo que muestra (a partir de la duda: "¿esto lo extrajiste del
# PDF o lo calculaste con GrupLAC?").
FUENTE_OFICIAL = ("#27ae60", "100% OFICIAL — extraído del PDF de MinCiencias (ScienTI)")
FUENTE_OFICIAL_MIXTA = (
    "#27ae60",
    "100% OFICIAL (PDF MinCiencias), salvo \"años de existencia\" que viene de GrupLAC",
)
FUENTE_GRUPLAC_CRUDO = (
    "#e67e22",
    "GRUPLAC — datos crudos, NO oficiales de la convocatoria (sin cuartil de revista). "
    "Solo para contraste visual con la pestaña anterior.",
)
FUENTE_CALCULADO = (
    "#2E86AB",
    "CALCULADO — proyección a partir del valor oficial + los productos que agregues aquí",
)


def _crear_badge_fuente(color: str, texto: str) -> QLabel:
    lbl = QLabel(texto)
    lbl.setWordWrap(True)
    lbl.setStyleSheet(
        f"background-color: {color}; color: white; font-weight: bold; "
        "font-size: 10px; padding: 4px 8px; border-radius: 3px;"
    )
    return lbl


class VistaSimulador957(QWidget):
    """Ficha de verificación y proyección Conv. 957 por grupo."""

    def __init__(self):
        super().__init__()
        self._sim = None
        self._error_carga = None
        try:
            self._sim = Simulador957()
        except Exception as e:
            self._error_carga = str(e)
        self._ajustes_pendientes: dict[str, int] = {}
        self._setup_ui()
        if self._sim is not None:
            self._poblar_combo_grupo()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(8, 8, 8, 8)

        titulo = QLabel("Simulador oficial Conv. 957 (MinCiencias / GrupLAC)")
        titulo.setFont(QFont("Arial", 14, QFont.Bold))
        titulo.setStyleSheet("color: #1a365d;")
        layout.addWidget(titulo)

        if self._error_carga:
            layout.addWidget(QLabel(
                f"No se pudo cargar medicion_957.xlsx: {self._error_carga}"
            ))
            return

        header = QHBoxLayout()
        header.addWidget(QLabel("Grupo (Conv. 957):"))
        self.combo_grupo = QComboBox()
        self.combo_grupo.setMinimumWidth(320)
        self.combo_grupo.currentTextChanged.connect(self._on_grupo_cambiado)
        header.addWidget(self.combo_grupo)

        self.lbl_badge_categoria = QLabel("")
        self.lbl_badge_categoria.setStyleSheet(
            "padding: 4px 12px; border-radius: 4px; color: white; font-weight: bold;"
        )
        header.addWidget(self.lbl_badge_categoria)
        header.addStretch()
        layout.addLayout(header)

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs, 1)

        self.tabs.addTab(self._crear_tab_resultado_oficial(), "Resultado oficial MinCiencias")
        self.tabs.addTab(self._crear_tab_productos_ventana(), "Productos y ventana (Paso 2)")
        self.tabs.addTab(self._crear_tab_gruplac(), "Productos en GrupLAC (auditoría)")
        self.tabs.addTab(self._crear_tab_proyeccion(), "Simulador de impacto")

    def _crear_tab_resultado_oficial(self) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        l.addWidget(_crear_badge_fuente(*FUENTE_OFICIAL_MIXTA))

        info_box = QGroupBox("Resumen del grupo")
        info_box.setStyleSheet("QGroupBox { font-weight: bold; color: #1a365d; }")
        info_lay = QHBoxLayout(info_box)
        self.lbl_area = QLabel("Área: —")
        self.lbl_ig_oficial = QLabel("IG oficial: —")
        self.lbl_cuartil_ig = QLabel("Cuartil IG: —")
        self.lbl_anios = QLabel("Años de existencia: —")
        for lbl in (self.lbl_area, self.lbl_ig_oficial, self.lbl_cuartil_ig, self.lbl_anios):
            lbl.setStyleSheet("font-size: 11px;")
            info_lay.addWidget(lbl)
        info_lay.addStretch()
        l.addWidget(info_box)

        self.tbl_indicadores = QTableWidget()
        cols = ["Indicador", "Valor oficial", "Máximo área", "Índice", "Ponderación", "Aporte a IG"]
        self.tbl_indicadores.setColumnCount(len(cols))
        self.tbl_indicadores.setHorizontalHeaderLabels(cols)
        self.tbl_indicadores.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.tbl_indicadores.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_indicadores.setAlternatingRowColors(True)
        l.addWidget(self.tbl_indicadores)
        return w

    def _crear_tab_productos_ventana(self) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        l.addWidget(_crear_badge_fuente(*FUENTE_OFICIAL))
        l.addWidget(QLabel(
            "Conteo oficial de productos por subtipo dentro de la ventana de "
            "observación (Paso 2 del proceso ScienTI) — tal cual lo calculó MinCiencias."
        ))
        self.tbl_productos = QTableWidget()
        cols = ["Sección", "Subtipo", "Total en ventana", "Ventana (años)", "λ"]
        self.tbl_productos.setColumnCount(len(cols))
        self.tbl_productos.setHorizontalHeaderLabels(cols)
        self.tbl_productos.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_productos.setAlternatingRowColors(True)
        l.addWidget(self.tbl_productos)
        return w

    def _crear_tab_gruplac(self) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        l.addWidget(_crear_badge_fuente(*FUENTE_GRUPLAC_CRUDO))
        self.lbl_gruplac_estado = QLabel("")
        self.lbl_gruplac_estado.setStyleSheet("color: #7f8c8d; font-size: 11px;")
        l.addWidget(self.lbl_gruplac_estado)

        self.tree_gruplac = QTreeWidget()
        self.tree_gruplac.setHeaderLabels(
            ["Categoría 957 / Subcategoría / Año / Producto", "Cantidad"]
        )
        self.tree_gruplac.setColumnWidth(0, 480)
        l.addWidget(self.tree_gruplac, 1)
        return w

    def _crear_tab_proyeccion(self) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        l.addWidget(_crear_badge_fuente(*FUENTE_CALCULADO))
        l.addWidget(QLabel(
            "Agrega productos hipotéticos (que el grupo aún no tiene) y observa el "
            "impacto proyectado en sus indicadores, el IG y la categoría."
        ))

        form = QHBoxLayout()
        form.addWidget(QLabel("Sección:"))
        self.combo_seccion = QComboBox()
        self.combo_seccion.addItems(SECCIONES_CON_PRODUCTOS)
        self.combo_seccion.currentTextChanged.connect(self._poblar_combo_subtipo)
        form.addWidget(self.combo_seccion)

        form.addWidget(QLabel("Subtipo:"))
        self.combo_subtipo = QComboBox()
        self.combo_subtipo.setMinimumWidth(160)
        form.addWidget(self.combo_subtipo)

        form.addWidget(QLabel("Unidades a agregar:"))
        self.spin_unidades = QSpinBox()
        self.spin_unidades.setRange(1, 999)
        form.addWidget(self.spin_unidades)

        btn_agregar = QPushButton("Agregar al escenario")
        btn_agregar.setStyleSheet(BTN_PRIMARIO)
        btn_agregar.clicked.connect(self._agregar_ajuste)
        form.addWidget(btn_agregar)
        form.addStretch()
        l.addLayout(form)

        fila_botones = QHBoxLayout()
        self.tbl_ajustes = QTableWidget()
        self.tbl_ajustes.setColumnCount(3)
        self.tbl_ajustes.setHorizontalHeaderLabels(["Subtipo", "+ Unidades", ""])
        self.tbl_ajustes.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_ajustes.setMaximumHeight(140)
        l.addWidget(self.tbl_ajustes)

        btn_calcular = QPushButton("Calcular proyección")
        btn_calcular.setStyleSheet(BTN_PRIMARIO)
        btn_calcular.clicked.connect(self._calcular_proyeccion)
        fila_botones.addWidget(btn_calcular)

        btn_limpiar = QPushButton("Limpiar escenario")
        btn_limpiar.setStyleSheet(BTN_SECUNDARIO)
        btn_limpiar.clicked.connect(self._limpiar_ajustes)
        fila_botones.addWidget(btn_limpiar)
        fila_botones.addStretch()
        l.addLayout(fila_botones)

        self.lbl_proyeccion_categoria = QLabel("")
        self.lbl_proyeccion_categoria.setStyleSheet(
            "font-size: 13px; font-weight: bold; padding: 6px 10px; "
            "background: #f0f0f0; border-radius: 4px;"
        )
        l.addWidget(self.lbl_proyeccion_categoria)

        self.tbl_proyeccion = QTableWidget()
        cols = ["Indicador", "Valor actual", "Valor proyectado", "Δ"]
        self.tbl_proyeccion.setColumnCount(len(cols))
        self.tbl_proyeccion.setHorizontalHeaderLabels(cols)
        self.tbl_proyeccion.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_proyeccion.setAlternatingRowColors(True)
        l.addWidget(self.tbl_proyeccion)

        l.addWidget(QLabel(
            "<i>Nota: no se evalúa la condición de \"investigador emérito/sénior/"
            "asociado vinculado\" — ese dato no existe en ningún export de GrupLAC. "
            "La categoría proyectada es un techo, no una garantía.</i>"
        ))
        return w

    # ------------------------------------------------------------------
    # Carga / cambio de grupo
    # ------------------------------------------------------------------

    def _poblar_combo_grupo(self):
        nombres = sorted(self._sim.df_grupos["nombre_grupo"].dropna().unique())
        self.combo_grupo.blockSignals(True)
        self.combo_grupo.addItems(nombres)
        self.combo_grupo.blockSignals(False)
        self._poblar_combo_subtipo(self.combo_seccion.currentText())
        if nombres:
            self._on_grupo_cambiado(self.combo_grupo.currentText())

    def _poblar_combo_subtipo(self, seccion: str):
        subtipos = sorted(
            s for s, sec in self._sim.seccion_de_subtipo.items() if sec == seccion
        )
        self.combo_subtipo.clear()
        self.combo_subtipo.addItems(subtipos)

    def _on_grupo_cambiado(self, nombre: str):
        if not nombre:
            return
        self._limpiar_ajustes()
        resultado = self._sim.simular(nombre)

        color = COLOR_CATEGORIA_957.get(resultado.categoria_oficial, "#7f8c8d")
        self.lbl_badge_categoria.setStyleSheet(
            f"padding: 4px 12px; border-radius: 4px; color: white; "
            f"font-weight: bold; background-color: {color};"
        )
        self.lbl_badge_categoria.setText(
            f"Categoría oficial: {resultado.categoria_oficial or 'sin clasificar'}"
        )

        self.lbl_area.setText(f"Área: {resultado.area or '—'}")
        ig_txt = f"{resultado.ig_oficial:.4f}" if resultado.ig_oficial else "—"
        self.lbl_ig_oficial.setText(f"IG oficial: {ig_txt}")
        self.lbl_cuartil_ig.setText(f"Cuartil IG: {resultado.cuartil_ig}")
        anios_txt = (
            f"{resultado.anios_existencia:.0f}"
            if resultado.anios_existencia is not None else "desconocido"
        )
        self.lbl_anios.setText(f"Años de existencia: {anios_txt}")

        self._poblar_tabla_indicadores(nombre, resultado)
        self._poblar_tabla_productos(nombre)
        self._poblar_tabla_gruplac(nombre)

    def _poblar_tabla_indicadores(self, nombre: str, resultado):
        orden = list(ETIQUETAS_INDICADOR)
        self.tbl_indicadores.setRowCount(len(orden))
        for fila, ind in enumerate(orden):
            valor = resultado.valores_simulados.get(ind, 0.0)
            indice = resultado.indices_simulados.get(ind, 0.0)
            pond = self._sim.ponderaciones.get(ind, 0.0)
            maximo = self._sim._maximo_oficial_por_grupo.get((nombre, ind))
            self.tbl_indicadores.setItem(fila, 0, _item(ETIQUETAS_INDICADOR[ind]))
            self.tbl_indicadores.setItem(fila, 1, _item(f"{valor:,.3f}"))
            self.tbl_indicadores.setItem(fila, 2, _item(f"{maximo:,.3f}" if maximo else "—"))
            self.tbl_indicadores.setItem(fila, 3, _item(f"{indice:.4f}"))
            self.tbl_indicadores.setItem(fila, 4, _item(f"{pond:.1f}"))
            self.tbl_indicadores.setItem(fila, 5, _item(f"{pond * indice:.4f}"))
        self.tbl_indicadores.resizeColumnsToContents()

    def _poblar_tabla_productos(self, nombre: str):
        df = self._sim.df_productos
        sub = df[df["grupo"] == nombre].sort_values(["seccion", "subtipo"])
        self.tbl_productos.setRowCount(len(sub))
        for fila, (_, r) in enumerate(sub.iterrows()):
            self.tbl_productos.setItem(fila, 0, _item(r["seccion"]))
            self.tbl_productos.setItem(fila, 1, _item(r["subtipo"]))
            self.tbl_productos.setItem(fila, 2, _item(int(r["total"])))
            self.tbl_productos.setItem(fila, 3, _item(int(r["ventana"])))
            self.tbl_productos.setItem(fila, 4, _item(f"{r['lambda_val']:.4f}"))
        self.tbl_productos.resizeColumnsToContents()

    def _poblar_tabla_gruplac(self, nombre: str):
        self.tree_gruplac.clear()
        nombre_gruplac = self._sim.nombre_en_gruplac(nombre)
        if nombre_gruplac is None:
            self.lbl_gruplac_estado.setText("No se encontró este grupo en GrupLAC.")
            return

        productos = self._sim.productos_gruplac(nombre)
        self.lbl_gruplac_estado.setText(
            f"Emparejado con \"{nombre_gruplac}\" en GrupLAC — {len(productos)} producto(s)."
        )

        # Agrupar Categoría 957 -> Subcategoría -> Año -> productos, para que
        # sea legible cuando hay cientos de filas (no es una lista plana).
        arbol: dict = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
        for p in productos:
            cat = p.get("categoria_957") or "Sin categoría"
            subcat = p.get("subcategoria_957") or "Sin subcategoría"
            anio = p.get("anio") or "Sin año"
            arbol[cat][subcat][anio].append(p)

        for cat in sorted(arbol):
            total_cat = sum(
                len(prods)
                for subcats in arbol[cat].values()
                for prods in subcats.values()
            )
            item_cat = QTreeWidgetItem([cat, str(total_cat)])
            color = COLORES_CATEGORIA.get(cat, "#7f8c8d")
            item_cat.setForeground(0, QColor(color))
            f = item_cat.font(0); f.setBold(True); item_cat.setFont(0, f)
            self.tree_gruplac.addTopLevelItem(item_cat)

            for subcat in sorted(arbol[cat]):
                total_subcat = sum(len(v) for v in arbol[cat][subcat].values())
                item_subcat = QTreeWidgetItem([subcat, str(total_subcat)])
                item_cat.addChild(item_subcat)

                for anio in sorted(arbol[cat][subcat], reverse=True):
                    prods_anio = arbol[cat][subcat][anio]
                    item_anio = QTreeWidgetItem([str(anio), str(len(prods_anio))])
                    item_subcat.addChild(item_anio)

                    for p in prods_anio:
                        tipo = p.get("tipo_producto_957") or ""
                        titulo = p.get("titulo") or ""
                        texto = f"{tipo}: {titulo}" if tipo else titulo
                        item_anio.addChild(QTreeWidgetItem([texto, ""]))

        self.tree_gruplac.collapseAll()

    # ------------------------------------------------------------------
    # Simulador de impacto (proyección)
    # ------------------------------------------------------------------

    def _agregar_ajuste(self):
        subtipo = self.combo_subtipo.currentText()
        if not subtipo:
            return
        n = self.spin_unidades.value()
        self._ajustes_pendientes[subtipo] = self._ajustes_pendientes.get(subtipo, 0) + n
        self._refrescar_tabla_ajustes()

    def _quitar_ajuste(self, subtipo: str):
        self._ajustes_pendientes.pop(subtipo, None)
        self._refrescar_tabla_ajustes()

    def _refrescar_tabla_ajustes(self):
        self.tbl_ajustes.setRowCount(len(self._ajustes_pendientes))
        for fila, (subtipo, n) in enumerate(self._ajustes_pendientes.items()):
            self.tbl_ajustes.setItem(fila, 0, _item(subtipo))
            self.tbl_ajustes.setItem(fila, 1, _item(f"+{n}"))
            btn = QPushButton("Quitar")
            btn.setStyleSheet(BTN_SECUNDARIO)
            btn.clicked.connect(lambda _checked, s=subtipo: self._quitar_ajuste(s))
            self.tbl_ajustes.setCellWidget(fila, 2, btn)
        self.tbl_ajustes.resizeColumnsToContents()

    def _limpiar_ajustes(self):
        self._ajustes_pendientes = {}
        self._refrescar_tabla_ajustes()
        self.tbl_proyeccion.setRowCount(0)
        self.lbl_proyeccion_categoria.setText("")

    def _calcular_proyeccion(self):
        nombre = self.combo_grupo.currentText()
        if not nombre:
            return
        if not self._ajustes_pendientes:
            QMessageBox.information(
                self, "Simulador de impacto",
                "Agrega al menos un producto al escenario antes de calcular.",
            )
            return
        try:
            resultado = self._sim.proyectar_productos(nombre, dict(self._ajustes_pendientes))
        except ValueError as e:
            QMessageBox.warning(self, "Simulador de impacto", str(e))
            return

        cat_actual = resultado.categoria_oficial or "—"
        cat_proy = resultado.categoria_simulada
        cambia = cat_actual != cat_proy
        color = COLOR_CATEGORIA_957.get(cat_proy, "#7f8c8d")
        texto = f"Categoría actual: {cat_actual}  →  Categoría proyectada: {cat_proy}"
        if cambia:
            texto += "  (¡cambia!)"
        self.lbl_proyeccion_categoria.setText(texto)
        self.lbl_proyeccion_categoria.setStyleSheet(
            "font-size: 13px; font-weight: bold; padding: 6px 10px; border-radius: 4px; "
            f"color: white; background-color: {color};"
        )

        orden = list(ETIQUETAS_INDICADOR)
        self.tbl_proyeccion.setRowCount(len(orden))
        valores_actuales = self._sim.simular(nombre).valores_simulados
        for fila, ind in enumerate(orden):
            actual = valores_actuales.get(ind, 0.0)
            proyectado = resultado.valores_simulados.get(ind, 0.0)
            delta = proyectado - actual
            self.tbl_proyeccion.setItem(fila, 0, _item(ETIQUETAS_INDICADOR[ind]))
            self.tbl_proyeccion.setItem(fila, 1, _item(f"{actual:,.3f}"))
            self.tbl_proyeccion.setItem(fila, 2, _item(f"{proyectado:,.3f}"))
            it_delta = _item(f"{delta:+,.3f}")
            if abs(delta) > 1e-6:
                it_delta.setForeground(QColor("#1a8f3c" if delta > 0 else "#c0392b"))
            self.tbl_proyeccion.setItem(fila, 3, it_delta)
        self.tbl_proyeccion.resizeColumnsToContents()
