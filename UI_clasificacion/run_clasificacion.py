"""
run_clasificacion.py

Punto de entrada standalone de UI_clasificacion -- separado de main_10.py a
propósito (se integrará a la app principal más adelante). Requiere que la
BD interna ya exista (data/db/academia_utp_integrado.db) y que haya al
menos una carpeta 'data/reporte excel_<fecha>' con el scrape de GrupLAC.

Uso:
    python UI_clasificacion/run_clasificacion.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))
sys.path.insert(0, str(Path(__file__).parent))

from PyQt5.QtWidgets import QApplication, QMainWindow

from main_10 import DatabaseManager
from vista_clasificacion import VistaClasificacion957


def main():
    app = QApplication(sys.argv)

    db = DatabaseManager()
    ventana = QMainWindow()
    ventana.setWindowTitle("Clasificación 957 (UI separada)")
    ventana.setGeometry(100, 100, 1400, 800)
    ventana.setCentralWidget(VistaClasificacion957(db))
    ventana.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
