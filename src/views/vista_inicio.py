"""
Vista de Inicio.
Pantalla de bienvenida: explica qué hace el software y permite cargar/actualizar
los archivos Excel de entrada. Solo se reprocesan los archivos que el usuario
suba de nuevo; los que ya estaban cargados se dejan como están.
"""

import json
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtWidgets import (
    QFileDialog, QGridLayout, QGroupBox, QHBoxLayout, QInputDialog, QLabel,
    QMessageBox, QPushButton, QScrollArea, QTableWidget, QTableWidgetItem,
    QVBoxLayout, QWidget,
)

from constants import (
    ARCHIVOS_FUENTE_957,
    COLUMNAS_CLAVE_POR_CATEGORIA,
    COLUMNAS_CONOCIDAS_POR_CATEGORIA,
)
from utils import normalizar_columna, obtener_directorio_base

# Hojas a inspeccionar por categoría cuando el Excel tiene varias (None = todas).
_HOJAS_POR_CATEGORIA = {
    "extension": ["Consolidado"],
    "cgt0104_2025": None,  # se filtran por nombre de hoja más abajo
    "cgt0104_2024": None,
}


class VistaInicio(QWidget):
    """Pantalla inicial: presentación del software + estado de datos de entrada."""

    procesar_solicitado = pyqtSignal()

    def __init__(self, db):
        super().__init__()
        self.db = db
        self.base_dir = obtener_directorio_base()
        self._labels_stats = {}
        self.setup_ui()
        self.actualizar_estado_archivos()

    # ------------------------------------------------------------------
    def setup_ui(self):
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)

        contenido = QWidget()
        scroll.setWidget(contenido)

        layout = QVBoxLayout(contenido)
        layout.setSpacing(16)
        layout.addWidget(self._panel_info())
        layout.addWidget(self._panel_archivos())
        layout.addWidget(self._panel_estadisticas())
        layout.addStretch()

        layout_principal = QVBoxLayout(self)
        layout_principal.setContentsMargins(0, 0, 0, 0)
        layout_principal.addWidget(scroll)

    # ------------------------------------------------------------------
    def _panel_info(self):
        box = QGroupBox("¿Qué es Consolidado de Información?")
        layout = QVBoxLayout(box)

        descripcion = QLabel(
            "Sistema de apoyo a la gestión de los grupos de investigación de la "
            "Universidad Tecnológica de Pereira (UTP). Consolida en una sola base de "
            "datos local la información que hoy vive repartida en varios archivos Excel "
            "institucionales: integrantes de grupos, producción académica, extensión, "
            "trabajos de grado, libros, innovación, proyectos y propiedad intelectual."
        )
        descripcion.setWordWrap(True)
        layout.addWidget(descripcion)

        modulos = QLabel(
            "<b>Qué se puede hacer en las otras pestañas:</b><br>"
            "&bull; <b>Búsqueda de Personas</b>: ver el consolidado completo de una persona "
            "(publicaciones, extensión, trabajos de grado, proyectos, etc.)<br>"
            "&bull; <b>Reportes por Grupo</b>: generar reportes en Excel y PDF por grupo de investigación.<br>"
            "&bull; <b>Seguimiento Grupos</b>: comparar lo cargado internamente contra GrupLAC y detectar "
            "productos pendientes de subir.<br>"
            "&bull; <b>Visor GrupLAC 957 / Simulador 957</b>: clasificar productos según la Convocatoria 957 "
            "de MinCiencias y simular el efecto de nuevos productos sobre la categoría del grupo."
        )
        modulos.setWordWrap(True)
        layout.addWidget(modulos)

        ayuda = QLabel(
            "<b>Cómo mantener los datos al día:</b> cuando llegue una versión nueva de alguno de los "
            "Excel, cárguela abajo en la fila correspondiente. Los datos que ya están cargados se dejan "
            "tal cual; el sistema solo reprocesa lo que usted actualice."
        )
        ayuda.setWordWrap(True)
        ayuda.setStyleSheet("color: #1a365d; padding-top: 6px;")
        layout.addWidget(ayuda)

        return box

    # ------------------------------------------------------------------
    def _panel_archivos(self):
        box = QGroupBox("Datos de entrada (Excel)")
        layout = QVBoxLayout(box)

        fila_agregar = QHBoxLayout()
        ayuda_agregar = QLabel(
            "Los archivos ya guardados se mantienen en memoria. Para añadir uno "
            "nuevo o actualizar uno existente, use el botón de abajo: el sistema "
            "detecta automáticamente a qué dato corresponde."
        )
        ayuda_agregar.setWordWrap(True)
        ayuda_agregar.setStyleSheet("color: #555;")
        btn_agregar = QPushButton("Agregar archivo…")
        btn_agregar.setStyleSheet(
            "background-color: #2e86ab; color: white; padding: 6px 14px; font-weight: bold;"
        )
        btn_agregar.clicked.connect(self._agregar_archivo)
        fila_agregar.addWidget(ayuda_agregar)
        fila_agregar.addWidget(btn_agregar)
        layout.addLayout(fila_agregar)

        self.tabla_archivos = QTableWidget()
        self.tabla_archivos.setColumnCount(4)
        self.tabla_archivos.setHorizontalHeaderLabels(
            ["Dato", "Archivo", "Última actualización", "Estado"]
        )
        self.tabla_archivos.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tabla_archivos.setSelectionMode(QTableWidget.NoSelection)
        self.tabla_archivos.verticalHeader().setVisible(False)
        layout.addWidget(self.tabla_archivos)

        fila_botones = QHBoxLayout()
        self.lbl_estado_carga = QLabel("")
        self.lbl_estado_carga.setStyleSheet("color: #555;")
        self.btn_procesar = QPushButton("Procesar datos nuevos")
        self.btn_procesar.setStyleSheet(
            "background-color: #1a365d; color: white; padding: 8px 16px; font-weight: bold;"
        )
        self.btn_procesar.clicked.connect(self.procesar_solicitado.emit)
        fila_botones.addWidget(self.lbl_estado_carga)
        fila_botones.addStretch()
        fila_botones.addWidget(self.btn_procesar)
        layout.addLayout(fila_botones)

        return box

    # ------------------------------------------------------------------
    def _panel_estadisticas(self):
        box = QGroupBox("Datos cargados actualmente")
        layout = QGridLayout(box)
        etiquetas = [
            ("personas", "Personas"), ("grupos", "Grupos"),
            ("publicaciones", "Publicaciones"), ("extensiones", "Extensiones"),
            ("trabajos", "Trabajos de grado"), ("innovacion", "Innovación"),
            ("proyectos", "Proyectos"), ("propiedad", "Propiedad intelectual"),
        ]
        for i, (clave, texto) in enumerate(etiquetas):
            valor = QLabel("0")
            valor.setFont(QFont("Arial", 18, QFont.Bold))
            valor.setStyleSheet("color: #1a365d;")
            valor.setAlignment(Qt.AlignCenter)
            nombre = QLabel(texto)
            nombre.setAlignment(Qt.AlignCenter)
            celda = QVBoxLayout()
            celda.addWidget(valor)
            celda.addWidget(nombre)
            cont = QWidget()
            cont.setLayout(celda)
            layout.addWidget(cont, i // 4, i % 4)
            self._labels_stats[clave] = valor
        return box

    # ------------------------------------------------------------------
    def _sello_carga(self):
        try:
            row = self.db.conn.execute(
                "SELECT valor FROM configuracion WHERE clave='sello_carga'"
            ).fetchone()
            return json.loads(row[0]) if row else {}
        except Exception:
            return {}

    def actualizar_estado_archivos(self):
        sello = self._sello_carga()
        self.tabla_archivos.setRowCount(len(ARCHIVOS_FUENTE_957))

        for fila, cat in enumerate(ARCHIVOS_FUENTE_957):
            ruta_encontrada = None
            for nombre in cat["variantes"]:
                ruta = self.base_dir / nombre
                if ruta.exists():
                    ruta_encontrada = ruta
                    break

            self.tabla_archivos.setItem(fila, 0, QTableWidgetItem(cat["label"]))

            if ruta_encontrada is None:
                self.tabla_archivos.setItem(fila, 1, QTableWidgetItem("—"))
                self.tabla_archivos.setItem(fila, 2, QTableWidgetItem("—"))
                estado_item = QTableWidgetItem("✗ No cargado")
                estado_item.setForeground(QColor("#a02020"))
            else:
                mtime = ruta_encontrada.stat().st_mtime
                self.tabla_archivos.setItem(fila, 1, QTableWidgetItem(ruta_encontrada.name))
                self.tabla_archivos.setItem(
                    fila, 2,
                    QTableWidgetItem(datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M")),
                )
                if ruta_encontrada.name == cat["variantes"][0]:
                    mtime_previo = sello.get(str(ruta_encontrada))
                    if mtime_previo is None or mtime > mtime_previo:
                        estado_item = QTableWidgetItem("⏳ Pendiente de procesar")
                        estado_item.setForeground(QColor("#b9770e"))
                    else:
                        estado_item = QTableWidgetItem("✓ Procesado")
                        estado_item.setForeground(QColor("#1e7e34"))
                else:
                    estado_item = QTableWidgetItem("✓ Cargado")
                    estado_item.setForeground(QColor("#1e7e34"))
            self.tabla_archivos.setItem(fila, 3, estado_item)

        self.tabla_archivos.resizeColumnsToContents()
        self.tabla_archivos.horizontalHeader().setStretchLastSection(True)

    def _agregar_archivo(self):
        """Punto único de carga: el usuario elige cualquier Excel y el sistema
        detecta a qué categoría de ARCHIVOS_FUENTE_957 corresponde por sus
        columnas. Los archivos ya guardados no se tocan; esto solo añade o
        reemplaza el de la categoría detectada."""
        ruta_origen, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar archivo a cargar", str(self.base_dir),
            "Archivos Excel (*.xlsx *.xls)",
        )
        if not ruta_origen:
            return

        categoria = self._resolver_categoria(ruta_origen)
        if categoria is None:
            return

        if not self._confirmar_columnas_nuevas(ruta_origen, categoria):
            return

        destino = self.base_dir / categoria["variantes"][0]
        try:
            shutil.copyfile(ruta_origen, destino)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo copiar el archivo:\n{e}")
            return

        QMessageBox.information(
            self, "Archivo cargado",
            f'"{Path(ruta_origen).name}" se guardó como dato de entrada para '
            f'"{categoria["label"]}".\n\nPulse "Procesar datos nuevos" para incorporarlo.',
        )
        self.actualizar_estado_archivos()

    def _resolver_categoria(self, ruta_origen):
        """Determina a qué categoría de ARCHIVOS_FUENTE_957 corresponde el
        archivo elegido, comparando sus columnas contra las columnas clave
        (distintivas) de cada categoría. Si no reconoce ninguna o coincide con
        varias, deja que el usuario la elija de una lista en vez de adivinar."""
        candidatas = []
        for cat in ARCHIVOS_FUENTE_957:
            columnas_clave = COLUMNAS_CLAVE_POR_CATEGORIA.get(cat["clave"])
            if not columnas_clave:
                continue
            detectadas = self._columnas_excel(ruta_origen, cat["clave"])
            if detectadas & columnas_clave:
                candidatas.append(cat)

        if len(candidatas) == 1:
            return candidatas[0]

        disponibles = candidatas or ARCHIVOS_FUENTE_957
        opciones = [c["label"] for c in disponibles]
        nombre_archivo = Path(ruta_origen).name
        if not candidatas:
            titulo = "Tipo de archivo no reconocido"
            texto = (
                f'No se pudo determinar automáticamente a qué dato corresponde '
                f'"{nombre_archivo}". Selecciónelo de la lista:'
            )
        else:
            titulo = "Tipo de archivo ambiguo"
            texto = (
                f'"{nombre_archivo}" coincide con más de un tipo de dato. '
                "Selecciónelo de la lista:"
            )

        elegido, ok = QInputDialog.getItem(self, titulo, texto, opciones, 0, False)
        if not ok:
            return None
        return next((c for c in disponibles if c["label"] == elegido), None)

    # ------------------------------------------------------------------
    # Detección de columnas nuevas (respecto a lo que el sistema ya sabe leer)
    # ------------------------------------------------------------------
    def _columnas_excel(self, ruta_origen, clave):
        """Devuelve el set de columnas (normalizadas) presentes en el Excel
        elegido, leyendo solo las hojas relevantes para esa categoría."""
        hojas = _HOJAS_POR_CATEGORIA.get(clave, [None])
        try:
            xls = pd.ExcelFile(ruta_origen, engine="openpyxl")
        except Exception:
            return set()

        if clave in ("cgt0104_2025", "cgt0104_2024"):
            hojas = [s for s in xls.sheet_names if "Soporte" in s or "Reg Propiedad" in s]
        elif hojas == [None]:
            hojas = [xls.sheet_names[0]]
        else:
            hojas = [s for s in hojas if s in xls.sheet_names]

        columnas = set()
        for hoja in hojas:
            try:
                encabezados = pd.read_excel(xls, sheet_name=hoja, nrows=0).columns
            except Exception:
                continue
            columnas.update(normalizar_columna(c) for c in encabezados)
        return columnas

    def _columnas_aceptadas(self, clave) -> set:
        try:
            row = self.db.conn.execute(
                "SELECT valor FROM configuracion WHERE clave=?",
                (f"columnas_aceptadas_{clave}",),
            ).fetchone()
            return set(json.loads(row[0])) if row else set()
        except Exception:
            return set()

    def _guardar_columnas_aceptadas(self, clave, columnas: set):
        self.db.conn.execute(
            "INSERT OR REPLACE INTO configuracion (clave, valor) VALUES (?, ?)",
            (f"columnas_aceptadas_{clave}", json.dumps(sorted(columnas))),
        )
        self.db.conn.commit()

    def _confirmar_columnas_nuevas(self, ruta_origen, categoria) -> bool:
        """
        Si la categoría tiene un set de columnas conocidas y el archivo elegido
        trae columnas que el sistema no reconoce, pregunta al usuario si desea
        continuar. Devuelve False si el usuario decide cancelar la carga.
        """
        clave = categoria["clave"]
        conocidas = COLUMNAS_CONOCIDAS_POR_CATEGORIA.get(clave)
        if conocidas is None:
            return True  # categoría sin columnas tabulares estables: no se verifica

        detectadas = self._columnas_excel(ruta_origen, clave)
        if not detectadas:
            return True  # no se pudo leer el encabezado; se deja procesar normalmente

        aceptadas = self._columnas_aceptadas(clave)
        nuevas = detectadas - conocidas - aceptadas

        # Columnas "clave" (discriminantes) de la categoría: columnas genéricas
        # como "nombre" o "email" aparecen en casi cualquier planilla, así que
        # por sí solas no bastan para considerar que el archivo es del tipo
        # correcto (ej. un archivo de notas también trae nombre/email).
        columnas_clave = COLUMNAS_CLAVE_POR_CATEGORIA.get(clave, conocidas)
        coincidentes_clave = detectadas & columnas_clave

        if not coincidentes_clave:
            # Ninguna columna distintiva está presente: probablemente es un
            # archivo de otro tipo de dato (formato distinto), no solo una
            # versión con un campo extra.
            mensaje = (
                f'El archivo elegido NO comparte ninguna columna distintiva con las que se '
                f'esperan para "{categoria["label"]}". Es probable que sea un archivo de otro '
                "tipo de dato (otro formato), no una versión nueva de este.\n\n"
                "Columnas que se esperaban (alguna de):\n"
                + "\n".join(f"  • {c}" for c in sorted(conocidas))
                + "\n\nColumnas encontradas en el archivo elegido:\n"
                + "\n".join(f"  • {c}" for c in sorted(detectadas))
                + "\n\n¿Desea cargarlo de todas formas?\n\n"
                "Si acepta: como ninguna columna se reconoce, todo el contenido del archivo "
                'se guardará como información adicional ("datos_adicionales") en vez de '
                "como datos normales, para no perderlo aunque no se use todavía.\n\n"
                "Si cancela: el archivo no se carga; revíselo y confirme que es el correcto "
                "para esta categoría."
            )
            respuesta = QMessageBox.question(
                self, "El archivo no coincide con el formato esperado", mensaje,
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if respuesta != QMessageBox.Yes:
                return False
            self._guardar_columnas_aceptadas(clave, aceptadas | nuevas)
            return True

        if not nuevas:
            return True

        mensaje = (
            f'El archivo para "{categoria["label"]}" trae columnas que el sistema todavía no '
            "reconoce, comparadas con la versión que ya procesa:\n\n"
            + "\n".join(f"  • {c}" for c in sorted(nuevas))
            + "\n\n¿Desea procesar esta información igual?\n\n"
            "Si acepta: las columnas conocidas se cargan como siempre y las columnas nuevas "
            'se guardan completas (sin perderlas) en el campo "datos_adicionales" de cada '
            "fila, aunque todavía no tengan una columna propia en los reportes.\n\n"
            "Si cancela: el archivo no se carga; podrá revisarlo y volver a intentarlo."
        )
        respuesta = QMessageBox.question(
            self, "Columnas nuevas detectadas", mensaje,
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes,
        )
        if respuesta != QMessageBox.Yes:
            return False

        self._guardar_columnas_aceptadas(clave, aceptadas | nuevas)
        return True

    # ------------------------------------------------------------------
    def actualizar_estadisticas(self, stats: dict):
        for clave, label in self._labels_stats.items():
            label.setText(str(stats.get(clave, 0)))

    def marcar_procesando(self, mensaje: str):
        self.lbl_estado_carga.setText(mensaje)

    def procesamiento_finalizado(self, stats: dict):
        self.actualizar_estadisticas(stats)
        self.actualizar_estado_archivos()
        self.lbl_estado_carga.setText("Datos al día.")
