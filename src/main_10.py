import json
import os
import re
import sqlite3
import sys
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
from pathlib import Path
from unidecode import unidecode
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDialog,
    QHeaderView, QHBoxLayout, QLabel, QLineEdit, QMainWindow,
    QMessageBox, QProgressBar, QPushButton, QSplitter,
    QTabWidget, QTableWidget, QTableWidgetItem, QTextEdit,
    QVBoxLayout, QWidget,
)

from analisis_seguimiento import AnalisisDuplicados
from views.vista_clasificacion_minciencias import VistaClasificacionMinCiencias
from views.vista_inicio import VistaInicio
from views.vista_seguimiento_grupos import VistaSeguimientoGrupos
from views.vista_simulador_957 import VistaSimulador957
from views.vista_visor_gruplac_957 import VisorGrupLAC957

from constants import COLUMNAS_CONOCIDAS_POR_CATEGORIA

warnings.filterwarnings("ignore")

from utils import (
    limpiar_cedula,
    limpiar_nombre_archivo,
    limpiar_texto,
    norm_text,
    normalizar_columna,
    normalizar_nombre,
    obtener_directorio_base,
)

# ==================== BASE DE DATOS ====================

class DatabaseManager:
    def __init__(self, db_path="data/db/academia_utp_integrado.db"):
        base_dir = obtener_directorio_base()
        db_file = base_dir / db_path
        self.db_path = str(db_file)
        if not db_file.exists():
            raise FileNotFoundError(f"No se encontró la BD: {db_file}")
        self.conn = sqlite3.connect(self.db_path, check_same_thread=False)
        self._aplicar_pragma()
        self.crear_tablas()
        self.mapa_cedulas = {}

    def _aplicar_pragma(self):
        """Configura SQLite para máximo rendimiento (seguro en lecturas y escrituras normales)."""
        self.conn.executescript("""
            PRAGMA journal_mode  = WAL;
            PRAGMA synchronous   = NORMAL;
            PRAGMA cache_size    = -32000;
            PRAGMA temp_store    = MEMORY;
            PRAGMA mmap_size     = 268435456;
            PRAGMA page_size     = 4096;
        """)

    
    def crear_tablas(self):
        cursor = self.conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS personas (
                cedula TEXT PRIMARY KEY,
                nombre TEXT NOT NULL,
                email TEXT,
                facultad TEXT,
                tipo_persona TEXT
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS cedulas_duplicadas (
                cedula_duplicada TEXT PRIMARY KEY,
                cedula_principal TEXT,
                FOREIGN KEY (cedula_principal) REFERENCES personas (cedula)
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS grupos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cedula TEXT,
                grupo TEXT,
                facultad TEXT,
                tipo_miembro TEXT,
                UNIQUE(cedula, grupo),
                FOREIGN KEY (cedula) REFERENCES personas (cedula)
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS publicaciones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cedula TEXT,
                titulo TEXT,
                revista_libro TEXT,
                doi_url TEXT,
                issn_isbn TEXT,
                año INTEGER,
                tipo TEXT,
                categoria TEXT,
                estado TEXT,
                grupo TEXT,
                fuente TEXT,
                FOREIGN KEY (cedula) REFERENCES personas (cedula)
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS extensiones (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cedula TEXT,
                actividad TEXT,
                tipo TEXT,
                modalidad TEXT,
                estado TEXT,
                fecha_inicio TEXT,
                fecha_fin TEXT,
                año INTEGER,
                poblacion TEXT,
                grupo TEXT,
                facultad TEXT,
                financiacion_interna TEXT,
                financiacion_externa TEXT,
                fuente TEXT,
                FOREIGN KEY (cedula) REFERENCES personas (cedula)
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS trabajos_grado (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cedula_director TEXT,
                nombre_director TEXT,
                cedula_estudiante TEXT,
                nombre_estudiante TEXT,
                titulo TEXT,
                programa TEXT,
                año INTEGER,
                estado TEXT,
                fecha_sustentacion TEXT,
                calificacion TEXT,
                facultad TEXT,
                fuente TEXT,
                FOREIGN KEY (cedula_director) REFERENCES personas (cedula)
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS productos_innovacion (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cedula TEXT,
                tipo_producto TEXT,
                nombre TEXT,
                descripcion TEXT,
                año INTEGER,
                estado TEXT,
                grupo TEXT,
                fuente TEXT,
                FOREIGN KEY (cedula) REFERENCES personas (cedula)
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS proyectos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cedula TEXT,
                responsable TEXT,
                titulo TEXT,
                objetivo TEXT,
                codigo_cie TEXT,
                tipo TEXT,
                año INTEGER,
                fecha_inicio TEXT,
                fecha_fin TEXT,
                estado TEXT,
                facultad TEXT,
                grupo TEXT,
                valor_aprobado TEXT,
                fuente TEXT,
                FOREIGN KEY (cedula) REFERENCES personas (cedula)
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS propiedad_intelectual (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                responsable TEXT,
                tipo_producto TEXT,
                tipo_patente TEXT,
                nombre_producto TEXT,
                numero_registro TEXT,
                proyecto TEXT,
                fecha_aprobacion TEXT,
                entidad TEXT,
                facultad TEXT,
                fuente TEXT
            )
        ''')
        
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS configuracion (
                clave TEXT PRIMARY KEY,
                valor TEXT
            )
        """)

        cursor.execute("CREATE INDEX IF NOT EXISTS idx_publicaciones_grupo ON publicaciones(grupo)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_extensiones_grupo ON extensiones(grupo)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_productos_innovacion_grupo ON productos_innovacion(grupo)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_proyectos_grupo ON proyectos(grupo)")

        self.conn.commit()
        self._migrar_columnas_adicionales()
        self._reparar_cedulas_duplicadas_huerfanas()

    def _reparar_cedulas_duplicadas_huerfanas(self):
        """
        Repara el daño de un bug histórico en fusionar_personas(): al volver a
        ejecutar la consolidación de duplicados sobre un par ya fusionado, el
        código resolvía cedula_dup a través del mapeo existente y terminaba
        igual a cedula_principal; el DELETE FROM personas final borraba
        entonces a la propia persona "principal" en cada recarga de datos.

        Por cada fila de cedulas_duplicadas cuyo cedula_principal ya no existe
        en personas:
          - Si cedula_duplicada SÍ existe en personas: es la identidad que
            sobrevivió. Se le reasignan los datos (grupos, publicaciones, etc.)
            que quedaron huérfanos bajo cedula_principal, y se borra el mapeo
            (ya no hay un duplicado real, solo una persona).
          - Si ninguna de las dos existe: no hay nada que mostrar todavía (los
            datos quedan huérfanos hasta que se vuelva a cargar el archivo
            fuente con esa cédula); solo se limpia el mapeo inválido.
        """
        cursor = self.conn.cursor()
        filas = cursor.execute(
            "SELECT cedula_duplicada, cedula_principal FROM cedulas_duplicadas"
        ).fetchall()

        tablas_cedula = ["grupos", "publicaciones", "extensiones",
                          "productos_innovacion", "proyectos"]
        reparados = 0
        for cedula_dup, cedula_principal in filas:
            existe_principal = cursor.execute(
                "SELECT 1 FROM personas WHERE cedula = ?", (cedula_principal,)
            ).fetchone()
            if existe_principal:
                continue  # mapeo sano, no tocar

            existe_dup = cursor.execute(
                "SELECT 1 FROM personas WHERE cedula = ?", (cedula_dup,)
            ).fetchone()
            if existe_dup and cedula_dup != cedula_principal:
                for tabla in tablas_cedula:
                    cursor.execute(
                        f"UPDATE OR IGNORE {tabla} SET cedula = ? WHERE cedula = ?",
                        (cedula_dup, cedula_principal),
                    )
                cursor.execute(
                    "UPDATE trabajos_grado SET cedula_director = ? WHERE cedula_director = ?",
                    (cedula_dup, cedula_principal),
                )
                cursor.execute(
                    "UPDATE trabajos_grado SET cedula_estudiante = ? WHERE cedula_estudiante = ?",
                    (cedula_dup, cedula_principal),
                )
                reparados += 1

            cursor.execute(
                "DELETE FROM cedulas_duplicadas WHERE cedula_duplicada = ? AND cedula_principal = ?",
                (cedula_dup, cedula_principal),
            )

        self.conn.commit()
        return reparados

    def _migrar_columnas_adicionales(self):
        """
        Agrega la columna 'datos_adicionales' (JSON con columnas de Excel no
        reconocidas) a las tablas alimentadas por archivo, si todavía no existe.
        CREATE TABLE IF NOT EXISTS no agrega columnas a tablas ya creadas en
        versiones anteriores de la BD, por eso se necesita este ALTER explícito.
        """
        tablas = ["publicaciones", "extensiones", "productos_innovacion", "propiedad_intelectual"]
        cursor = self.conn.cursor()
        for tabla in tablas:
            columnas = {fila[1] for fila in cursor.execute(f"PRAGMA table_info({tabla})")}
            if "datos_adicionales" not in columnas:
                cursor.execute(f"ALTER TABLE {tabla} ADD COLUMN datos_adicionales TEXT")
        self.conn.commit()
        self._migrar_indices_unicos()

    # Columnas que identifican una fila como "la misma" entre cargas (para que
    # al volver a cargar un archivo se actualice/acumule en vez de duplicarse).
    _CLAVES_DEDUP = {
        "publicaciones": ("cedula", "titulo", "año"),
        "extensiones": ("cedula", "actividad", "fecha_inicio"),
        "productos_innovacion": ("cedula", "nombre", "año", "grupo"),
        "proyectos": ("cedula", "titulo", "año"),
        "propiedad_intelectual": ("responsable", "nombre_producto", "numero_registro"),
        "trabajos_grado": ("cedula_estudiante", "titulo"),
    }

    def _migrar_indices_unicos(self):
        """
        Crea un índice único por tabla (ver _CLAVES_DEDUP) para poder usar
        INSERT OR REPLACE como mecanismo de "actualizar si existe, agregar si
        es nuevo": al recargar un archivo, los registros ya guardados nunca se
        borran; solo se actualizan si cambiaron y se acumulan los nuevos.
        Antes de crear el índice se eliminan duplicados exactos preexistentes
        (misma combinación de columnas clave), conservando la fila de menor id.
        """
        cursor = self.conn.cursor()
        for tabla, columnas in self._CLAVES_DEDUP.items():
            col_list = ", ".join(columnas)
            cursor.execute(f"""
                DELETE FROM {tabla}
                WHERE id NOT IN (SELECT MIN(id) FROM {tabla} GROUP BY {col_list})
            """)
            cursor.execute(
                f"CREATE UNIQUE INDEX IF NOT EXISTS ux_{tabla}_dedup ON {tabla}({col_list})"
            )
        self.conn.commit()

    # ── Caché de carga ───────────────────────────────────────────────────────

    def guardar_sello_carga(self, archivos: list):
        """
        Guarda los timestamps de los archivos procesados para detectar cambios.
        archivos: lista de rutas (str o Path).
        """
        import json
        from pathlib import Path as _P
        sello = {}
        for ruta in archivos:
            p = _P(ruta)
            if p.exists():
                sello[str(p)] = p.stat().st_mtime
        self.conn.execute(
            "INSERT OR REPLACE INTO configuracion (clave, valor) VALUES ('sello_carga', ?)",
            (json.dumps(sello),),
        )
        self.conn.commit()

    def cache_valida(self, archivos_a_verificar: list) -> bool:
        """
        Devuelve True si la BD ya tiene datos y ningún archivo fuente
        fue modificado desde la última carga completa.
        """
        import json
        from pathlib import Path as _P
        try:
            row = self.conn.execute(
                "SELECT valor FROM configuracion WHERE clave='sello_carga'"
            ).fetchone()
            if not row:
                return False
            n_personas = self.conn.execute(
                "SELECT COUNT(*) FROM personas"
            ).fetchone()[0]
            if n_personas == 0:
                return False
            sello = json.loads(row[0])
            for ruta in archivos_a_verificar:
                p = _P(ruta)
                if not p.exists():
                    continue
                mtime_actual = p.stat().st_mtime
                mtime_previo = sello.get(str(p))
                if mtime_previo is None or mtime_actual > mtime_previo:
                    return False
            return True
        except Exception:
            return False

    def obtener_cedula_principal(self, cedula):
        if cedula in self.mapa_cedulas:
            return self.mapa_cedulas[cedula]
        
        cursor = self.conn.cursor()
        resultado = cursor.execute(
            'SELECT cedula_principal FROM cedulas_duplicadas WHERE cedula_duplicada = ?',
            (cedula,)
        ).fetchone()
        
        if resultado:
            self.mapa_cedulas[cedula] = resultado[0]
            return resultado[0]
        
        return cedula

    def son_cedulas_duplicadas(self, cedula1, cedula2):
        if cedula1 == cedula2:
            return False
        if cedula1 + "0" == cedula2:
            return True
        if cedula2 + "0" == cedula1:
            return True
        return False
    
    def consolidar_duplicados(self):
        cursor = self.conn.cursor()
        personas = cursor.execute('SELECT cedula, nombre FROM personas ORDER BY nombre').fetchall()

        nombre_por_cedula = {c: n for c, n in personas}

        cedulas_por_nombre = {}
        for cedula, nombre in personas:
            nn = normalizar_nombre(nombre)
            cedulas_por_nombre.setdefault(nn, []).append(cedula)

        pairs = cursor.execute("""
            SELECT DISTINCT a.cedula AS c1, b.cedula AS c2
            FROM personas a
            JOIN personas b ON (b.cedula = a.cedula || '0' OR a.cedula = b.cedula || '0')
            WHERE a.cedula < b.cedula
        """).fetchall()

        adj = {}
        for c1, c2 in pairs:
            adj.setdefault(c1, set()).add(c2)
            adj.setdefault(c2, set()).add(c1)

        duplicados_encontrados = []
        global_processed = set()

        for nombre_norm, grupo_cedulas in cedulas_por_nombre.items():
            if len(grupo_cedulas) < 2:
                continue

            grupo_set = set(grupo_cedulas)

            for cedula in grupo_cedulas:
                if cedula in global_processed:
                    continue

                cluster = set()
                stack = [cedula]
                while stack:
                    cur = stack.pop()
                    if cur in global_processed or cur not in grupo_set:
                        continue
                    global_processed.add(cur)
                    cluster.add(cur)
                    for nb in adj.get(cur, set()):
                        if nb not in global_processed and nb in grupo_set:
                            stack.append(nb)

                if len(cluster) > 1:
                    cedula_principal = min(cluster, key=len)
                    cedulas_secundarias = [c for c in cluster if c != cedula_principal]
                    nombre_principal = nombre_por_cedula[cedula_principal]

                    self.fusionar_personas(cedula_principal, cedulas_secundarias, nombre_principal)

                    duplicados_encontrados.append({
                        'nombre': nombre_principal,
                        'cedula_principal': cedula_principal,
                        'cedulas': list(cluster),
                    })

        return duplicados_encontrados

    def fusionar_personas(self, cedula_principal, cedulas_duplicadas, nombre_principal):
        cursor = self.conn.cursor()
        
        try:
            # Actualizar nombre de la persona principal
            cursor.execute(
                'UPDATE personas SET nombre = ? WHERE cedula = ?',
                (nombre_principal, cedula_principal)
            )
            
            for cedula_dup in cedulas_duplicadas:
                cedula_dup = self.obtener_cedula_principal(cedula_dup)
                if cedula_dup == cedula_principal:
                    # Ya estaba fusionado en una corrida anterior: si se
                    # sigue de largo, las instrucciones de abajo (en
                    # particular el DELETE FROM personas) borrarían a la
                    # propia persona principal por error.
                    continue

                # Registrar duplicado
                cursor.execute('''
                    INSERT OR REPLACE INTO cedulas_duplicadas (cedula_duplicada, cedula_principal)
                    VALUES (?, ?)
                ''', (cedula_dup, cedula_principal))
                
                self.mapa_cedulas[cedula_dup] = cedula_principal
                
                # Mover datos - usar OR IGNORE para evitar conflictos de UNIQUE
                for tabla in ['grupos', 'publicaciones', 'extensiones', 'productos_innovacion', 'proyectos']:
                    try:
                        cursor.execute(f'UPDATE OR IGNORE {tabla} SET cedula = ? WHERE cedula = ?', 
                                     (cedula_principal, cedula_dup))
                    except:
                        pass
                
                # Trabajos de grado
                try:
                    cursor.execute('UPDATE trabajos_grado SET cedula_director = ? WHERE cedula_director = ?', 
                                 (cedula_principal, cedula_dup))
                except:
                    pass
                    
                try:
                    cursor.execute('UPDATE trabajos_grado SET cedula_estudiante = ? WHERE cedula_estudiante = ?', 
                                 (cedula_principal, cedula_dup))
                except:
                    pass
                
                # Copiar email y facultad si están vacíos en la principal
                persona_dup = cursor.execute(
                    'SELECT email, facultad FROM personas WHERE cedula = ?',
                    (cedula_dup,)
                ).fetchone()
                
                if persona_dup:
                    email_dup, facultad_dup = persona_dup
                    cursor.execute('''
                        UPDATE personas 
                        SET email = COALESCE(NULLIF(email, ''), ?),
                            facultad = COALESCE(NULLIF(facultad, ''), ?)
                        WHERE cedula = ?
                    ''', (email_dup, facultad_dup, cedula_principal))
                
                # Eliminar persona duplicada
                cursor.execute('DELETE FROM personas WHERE cedula = ?', (cedula_dup,))
            
            # Hacer commit una sola vez al final
            self.conn.commit()
            
        except sqlite3.OperationalError as e:
            # Si hay error, hacer rollback
            try:
                self.conn.rollback()
            except:
                pass
            # Re-raise el error para que se maneje arriba
            raise e
        except Exception as e:
            # Cualquier otro error
            try:
                self.conn.rollback()
            except:
                pass
            raise e

    def limpiar_datos(self):
        cursor = self.conn.cursor()
        tablas = ['grupos', 'publicaciones', 'extensiones', 'trabajos_grado', 
                  'productos_innovacion', 'proyectos', 'propiedad_intelectual', 
                  'cedulas_duplicadas', 'personas']
        for tabla in tablas:
            cursor.execute(f'DELETE FROM {tabla}')
        self.conn.commit()

    def insertar_persona(self, cedula, nombre, email="", facultad="", tipo=""):
        """Upsert de persona: INSERT si no existe, UPDATE campos vacíos si ya existe."""
        if not cedula:
            return
        cedula = self.obtener_cedula_principal(cedula)
        self.conn.execute(
            """
            INSERT INTO personas (cedula, nombre, email, facultad, tipo_persona)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(cedula) DO UPDATE SET
                email        = COALESCE(NULLIF(excluded.email, ''),        personas.email),
                facultad     = COALESCE(NULLIF(excluded.facultad, ''),     personas.facultad),
                tipo_persona = COALESCE(NULLIF(excluded.tipo_persona, ''), personas.tipo_persona)
            """,
            (cedula, nombre or "", email or "", facultad or "", tipo or ""),
        )
        # No se hace commit aquí; el llamador decide cuándo confirmar.

    def _upsert_personas_batch(self, filas):
        """Batch upsert de personas: [(cedula, nombre, email, facultad, tipo), ...]."""
        if not filas:
            return
        self.conn.executemany(
            """
            INSERT INTO personas (cedula, nombre, email, facultad, tipo_persona)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(cedula) DO UPDATE SET
                email        = COALESCE(NULLIF(excluded.email, ''),        personas.email),
                facultad     = COALESCE(NULLIF(excluded.facultad, ''),     personas.facultad),
                tipo_persona = COALESCE(NULLIF(excluded.tipo_persona, ''), personas.tipo_persona)
            """,
            filas,
        )

    def buscar_personas(self, termino):
        cursor = self.conn.cursor()
        termino_limpio = termino.strip()
        
        if termino_limpio.replace('.', '').replace(',', '').isdigit():
            cedula_limpia = limpiar_cedula(termino_limpio)
            query = '''
                SELECT cedula, nombre, email, facultad, tipo_persona
                FROM personas WHERE cedula LIKE ?
                ORDER BY nombre
            '''
            params = (f'%{cedula_limpia}%',)
        else:
            palabras = termino_limpio.split()
            if len(palabras) == 1:
                query = '''
                    SELECT cedula, nombre, email, facultad, tipo_persona
                    FROM personas WHERE LOWER(nombre) LIKE LOWER(?)
                    ORDER BY nombre
                '''
                params = (f'%{palabras[0]}%',)
            else:
                condiciones = ' AND '.join(['LOWER(nombre) LIKE LOWER(?)' for _ in palabras])
                query = f'''
                    SELECT cedula, nombre, email, facultad, tipo_persona
                    FROM personas WHERE {condiciones}
                    ORDER BY nombre
                '''
                params = tuple([f'%{p}%' for p in palabras])
        
        resultados = cursor.execute(query, params).fetchall()
        
        resultados_filtrados = []
        for resultado in resultados:
            cedula = resultado[0]
            # Solo se excluye si el "principal" al que apunta sigue existiendo
            # de verdad en personas; si no, mostrar esta fila es lo correcto
            # (evita ocultar personas por un mapeo de duplicado roto/huérfano).
            es_duplicada_valida = cursor.execute('''
                SELECT 1 FROM cedulas_duplicadas cd
                JOIN personas pp ON pp.cedula = cd.cedula_principal
                WHERE cd.cedula_duplicada = ?
            ''', (cedula,)).fetchone()

            if not es_duplicada_valida:
                resultados_filtrados.append(resultado)

        return resultados_filtrados

    def obtener_detalle_persona(self, cedula):
        cursor = self.conn.cursor()
        detalle = {}
        
        persona = cursor.execute(
            'SELECT * FROM personas WHERE cedula = ?', (cedula,)
        ).fetchone()
        
        if persona:
            detalle['info'] = {
                'cedula': persona[0],
                'nombre': persona[1],
                'email': persona[2] or 'No disponible',
                'facultad': persona[3] or 'No disponible',
                'tipo': persona[4] or 'No especificado'
            }
        
        detalle['grupos'] = cursor.execute(
            'SELECT * FROM grupos WHERE cedula = ? ORDER BY grupo', (cedula,)
        ).fetchall()
        
        detalle['publicaciones'] = cursor.execute(
            'SELECT * FROM publicaciones WHERE cedula = ? ORDER BY año DESC', (cedula,)
        ).fetchall()
        
        detalle['extensiones'] = cursor.execute(
            'SELECT * FROM extensiones WHERE cedula = ? ORDER BY año DESC', (cedula,)
        ).fetchall()
        
        detalle['trabajos_grado_director'] = cursor.execute(
            'SELECT * FROM trabajos_grado WHERE cedula_director = ? ORDER BY año DESC', (cedula,)
        ).fetchall()
        
        detalle['trabajos_grado_estudiante'] = cursor.execute(
            'SELECT * FROM trabajos_grado WHERE cedula_estudiante = ? ORDER BY año DESC', (cedula,)
        ).fetchall()
        
        detalle['productos_innovacion'] = cursor.execute(
            'SELECT * FROM productos_innovacion WHERE cedula = ? ORDER BY año DESC', (cedula,)
        ).fetchall()
        
        detalle['proyectos'] = cursor.execute(
            'SELECT * FROM proyectos WHERE cedula = ? ORDER BY año DESC', (cedula,)
        ).fetchall()
        
        return detalle

    def obtener_grupos(self):
        cursor = self.conn.cursor()
        return cursor.execute('''
            SELECT DISTINCT grupo FROM grupos 
            WHERE grupo IS NOT NULL AND grupo != ''
            ORDER BY grupo
        ''').fetchall()

    def obtener_integrantes_grupo(self, grupo):
        cursor = self.conn.cursor()
        return cursor.execute('''
            SELECT DISTINCT p.cedula, p.nombre, g.tipo_miembro, p.email, p.facultad
            FROM grupos g
            JOIN personas p ON g.cedula = p.cedula
            WHERE g.grupo = ?
            ORDER BY p.nombre
        ''', (grupo,)).fetchall()

    def obtener_productos_grupo(self, grupo, filtros=None):
        if filtros is None:
            filtros = ['Publicaciones', 'Extensiones', 'Trabajos de Grado', 'Productos Innovación', 'Proyectos']
        
        cursor = self.conn.cursor()
        cedulas = cursor.execute(
            'SELECT DISTINCT cedula FROM grupos WHERE grupo = ?', (grupo,)
        ).fetchall()
        
        if not cedulas:
            return []
        
        cedulas_list = [c[0] for c in cedulas]
        placeholders = ','.join(['?' for _ in cedulas_list])
        productos = []
        
        if 'Publicaciones' in filtros:
            query = f'''
                SELECT p.cedula, per.nombre, p.titulo, p.año, 'Publicación',
                       p.tipo, p.categoria, p.revista_libro, p.fuente
                FROM publicaciones p
                JOIN personas per ON p.cedula = per.cedula
                WHERE p.cedula IN ({placeholders})
                ORDER BY p.año DESC
            '''
            productos.extend(cursor.execute(query, cedulas_list).fetchall())
        
        if 'Extensiones' in filtros:
            query = f'''
                SELECT e.cedula, per.nombre, e.actividad, e.año, 'Extensión',
                       e.tipo, e.modalidad, e.estado, e.fuente
                FROM extensiones e
                JOIN personas per ON e.cedula = per.cedula
                WHERE e.cedula IN ({placeholders})
                ORDER BY e.año DESC
            '''
            productos.extend(cursor.execute(query, cedulas_list).fetchall())
        
        if 'Trabajos de Grado' in filtros:
            query = f'''
                SELECT t.cedula_director, per.nombre, t.titulo, t.año, 'Trabajo de Grado',
                       t.programa, t.estado, t.nombre_estudiante, t.fuente
                FROM trabajos_grado t
                JOIN personas per ON t.cedula_director = per.cedula
                WHERE t.cedula_director IN ({placeholders})
                ORDER BY t.año DESC
            '''
            productos.extend(cursor.execute(query, cedulas_list).fetchall())
        
        if 'Productos Innovación' in filtros:
            query = f'''
                SELECT pi.cedula, per.nombre, pi.nombre, pi.año, 'Innovación',
                       pi.tipo_producto, pi.estado, '', pi.fuente
                FROM productos_innovacion pi
                JOIN personas per ON pi.cedula = per.cedula
                WHERE pi.cedula IN ({placeholders})
                ORDER BY pi.año DESC
            '''
            productos.extend(cursor.execute(query, cedulas_list).fetchall())
        
        if 'Proyectos' in filtros:
            query = f'''
                SELECT pr.cedula, per.nombre, pr.titulo, pr.año, 'Proyecto',
                       pr.tipo, pr.estado, pr.codigo_cie, pr.fuente
                FROM proyectos pr
                JOIN personas per ON pr.cedula = per.cedula
                WHERE pr.cedula IN ({placeholders})
                ORDER BY pr.año DESC
            '''
            productos.extend(cursor.execute(query, cedulas_list).fetchall())
        
        return productos
    
    def obtener_productos_grupo_detallado(self, grupo, filtros=None):
        """Versión detallada que retorna diccionarios con toda la información"""
        if filtros is None:
            filtros = ['Publicaciones', 'Extensiones', 'Trabajos de Grado', 'Productos Innovación', 'Proyectos']
        
        cursor = self.conn.cursor()
        cedulas = cursor.execute(
            'SELECT DISTINCT cedula FROM grupos WHERE grupo = ?', (grupo,)
        ).fetchall()
        
        if not cedulas:
            return []
        
        cedulas_list = [c[0] for c in cedulas]
        placeholders = ','.join(['?' for _ in cedulas_list])
        productos = []
        
        if 'Publicaciones' in filtros:
            query = f'''
                SELECT p.*, per.nombre as nombre_investigador
                FROM publicaciones p
                JOIN personas per ON p.cedula = per.cedula
                WHERE p.cedula IN ({placeholders})
                ORDER BY p.año DESC
            '''
            rows = cursor.execute(query, cedulas_list).fetchall()
            for row in rows:
                productos.append({
                    'cedula': row[1],
                    'investigador': row[-1],
                    'titulo': row[2],
                    'revista_libro': row[3],
                    'doi_url': row[4],
                    'issn_isbn': row[5],
                    'año': row[6],
                    'tipo': row[7],
                    'categoria': row[8],
                    'estado': row[9],
                    'grupo': row[10],
                    'fuente': row[11],
                    'tipo_producto': 'Publicación',
                    'detalle': row[3] or ''
                })
        
        if 'Extensiones' in filtros:
            query = f'''
                SELECT e.*, per.nombre as nombre_investigador
                FROM extensiones e
                JOIN personas per ON e.cedula = per.cedula
                WHERE e.cedula IN ({placeholders})
                ORDER BY e.año DESC
            '''
            rows = cursor.execute(query, cedulas_list).fetchall()
            for row in rows:
                productos.append({
                    'cedula': row[1],
                    'investigador': row[-1],
                    'titulo': row[2],
                    'tipo': row[3],
                    'modalidad': row[4],
                    'estado': row[5],
                    'fecha_inicio': row[6],
                    'fecha_fin': row[7],
                    'año': row[8],
                    'poblacion': row[9],
                    'grupo': row[10],
                    'facultad': row[11],
                    'financiacion_interna': row[12],
                    'financiacion_externa': row[13],
                    'fuente': row[14],
                    'tipo_producto': 'Extensión',
                    'categoria': row[3] or '',
                    'detalle': row[4] or ''
                })
        
        if 'Trabajos de Grado' in filtros:
            query = f'''
                SELECT t.*, per.nombre as nombre_director
                FROM trabajos_grado t
                JOIN personas per ON t.cedula_director = per.cedula
                WHERE t.cedula_director IN ({placeholders})
                ORDER BY t.año DESC
            '''
            rows = cursor.execute(query, cedulas_list).fetchall()
            for row in rows:
                productos.append({
                    'cedula': row[1],
                    'investigador': row[-1],
                    'titulo': row[5],
                    'programa': row[6],
                    'año': row[7],
                    'estado': row[8],
                    'fecha_sustentacion': row[9],
                    'calificacion': row[10],
                    'facultad': row[11],
                    'estudiante': row[4],
                    'cedula_estudiante': row[3],
                    'fuente': row[12],
                    'tipo_producto': 'Trabajo de Grado',
                    'categoria': row[6] or '',
                    'detalle': row[4] or ''
                })
        
        if 'Productos Innovación' in filtros:
            query = f'''
                SELECT pi.*, per.nombre as nombre_investigador
                FROM productos_innovacion pi
                JOIN personas per ON pi.cedula = per.cedula
                WHERE pi.cedula IN ({placeholders})
                ORDER BY pi.año DESC
            '''
            rows = cursor.execute(query, cedulas_list).fetchall()
            for row in rows:
                productos.append({
                    'cedula': row[1],
                    'investigador': row[-1],
                    'titulo': row[3],
                    'tipo_producto_detalle': row[2],
                    'descripcion': row[4],
                    'año': row[5],
                    'estado': row[6],
                    'grupo': row[7],
                    'fuente': row[8],
                    'tipo_producto': 'Innovación',
                    'categoria': row[2] or '',
                    'detalle': row[4][:50] + '...' if row[4] and len(row[4]) > 50 else row[4] or ''
                })
        
        if 'Proyectos' in filtros:
            query = f'''
                SELECT pr.*, per.nombre as nombre_investigador
                FROM proyectos pr
                JOIN personas per ON pr.cedula = per.cedula
                WHERE pr.cedula IN ({placeholders})
                ORDER BY pr.año DESC
            '''
            rows = cursor.execute(query, cedulas_list).fetchall()
            for row in rows:
                productos.append({
                    'cedula': row[1],
                    'investigador': row[-1],
                    'responsable': row[2],
                    'titulo': row[3],
                    'objetivo': row[4],
                    'codigo_cie': row[5],
                    'tipo': row[6],
                    'año': row[7],
                    'fecha_inicio': row[8],
                    'fecha_fin': row[9],
                    'estado': row[10],
                    'facultad': row[11],
                    'grupo': row[12],
                    'valor_aprobado': row[13],
                    'fuente': row[14],
                    'tipo_producto': 'Proyecto',
                    'categoria': row[6] or '',
                    'detalle': row[5] or ''
                })
        
        return productos

    def obtener_estadisticas(self):
        """Obtiene conteos en una sola consulta multi-tabla."""
        row = self.conn.execute("""
            SELECT
                (SELECT COUNT(*)          FROM personas)            AS personas,
                (SELECT COUNT(DISTINCT grupo) FROM grupos)          AS grupos,
                (SELECT COUNT(*)          FROM publicaciones)       AS publicaciones,
                (SELECT COUNT(*)          FROM extensiones)         AS extensiones,
                (SELECT COUNT(*)          FROM trabajos_grado)      AS trabajos,
                (SELECT COUNT(*)          FROM productos_innovacion) AS innovacion,
                (SELECT COUNT(*)          FROM proyectos)           AS proyectos,
                (SELECT COUNT(*)          FROM propiedad_intelectual) AS propiedad
        """).fetchone()
        keys = ["personas", "grupos", "publicaciones", "extensiones",
                "trabajos", "innovacion", "proyectos", "propiedad"]
        return dict(zip(keys, row))
        
    def close(self):
        """Cierra correctamente la conexión SQLite"""
        try:
            self.conn.commit()
            self.conn.close()
        except Exception as e:
            print("Error cerrando BD:", e)       


# ==================== CARGADOR DE DATOS INTEGRADO ====================

class CargadorDatosIntegrado(QThread):
    progreso = pyqtSignal(str)
    finalizado = pyqtSignal(dict)
    duplicados_consolidados = pyqtSignal(list)
    
    def __init__(self, db):
        super().__init__()
        self.db = db
        # self.archivos_directorio = Path(".") Esta estaba anteriormente.
        self.archivos_directorio = obtener_directorio_base()
    

    def run(self):
        base = self.archivos_directorio
        nomina_nombres = [
            "Listado Integrantes Grupos de Investigación UTP  080825.xlsx",
            "Listado Integrantes Grupos de Investigación UTP - 080825.xlsx",
            "Listado Integrantes Grupos de Investigacion UTP  080825.xlsx",
            "Listado Integrantes Grupos de Investigacion UTP - 080825.xlsx",
            "Consolidado Extensión 2024.xlsx",
            "BASE DATOS PRODUCCIÓN 2024.xlsx",
            "BASE DATOS PRODUCCIÓN  2025  CIARP.xlsx",
            "Trabajos Grado  Trabajo de Grado.xlsx",
            "TrabajosGrado_TrabajoDeGrado 2024.xlsx",
            "Reporte_libros.xlsx",
            "info_productos_innovacion.xlsx",
            "Proyectos de investigación registrados en 2024.xlsx",
            "CGT0104 - No de productos resultados de investigacion 31072025.xlsx",
            "CGT0104 - No de productos resultados de investigacion 31122024.xlsx",
        ]
        archivos_fuente = [str(base / n) for n in nomina_nombres]

        if self.db.cache_valida(archivos_fuente):
            self.progreso.emit("Base de datos en caché — omitiendo reproceso")
            self.finalizado.emit(self.db.obtener_estadisticas())
            return

        conn = self.db.conn
        conn.execute("PRAGMA synchronous=OFF")
        try:
            if self.db.obtener_estadisticas()['personas'] == 0:
                # Primera carga sobre una BD vacía: no hay nada que conservar.
                self.db.limpiar_datos()
            conn.commit()
            # En cualquier otro caso NO se borra nada: los extractores leen de
            # nuevo todos los archivos presentes y los índices únicos por
            # categoría (DatabaseManager._CLAVES_DEDUP) hacen que insertar lo
            # mismo otra vez actualice la fila existente en vez de duplicarla,
            # mientras que las filas realmente nuevas simplemente se acumulan.

            # ── Extracción paralela ──────────────────────────────────────────
            self.progreso.emit("Extrayendo datos de archivos fuente…")
            extractores = {
                'integrantes':          self._extraer_integrantes,
                'extensiones':          self._extraer_extensiones,
                'produccion':           self._extraer_produccion,
                'trabajos_grado':       self._extraer_trabajos_grado,
                'libros':               self._extraer_libros,
                'productos_innovacion': self._extraer_productos_innovacion,
            }
            resultados = {}
            with ThreadPoolExecutor(max_workers=4) as pool:
                futuros = {pool.submit(fn): name for name, fn in extractores.items()}
                for futuro in as_completed(futuros):
                    name = futuros[futuro]
                    try:
                        resultados[name] = futuro.result()
                        self.progreso.emit(f"✓ Extraídos datos de {name}")
                    except Exception as e:
                        self.progreso.emit(f"Error extrayendo {name}: {e}")
                        resultados[name] = ([], [])

            # ── Inserción secuencial ─────────────────────────────────────────
            self.progreso.emit("Insertando datos en la base…")

            for name, (pers, _) in resultados.items():
                if pers:
                    self.db._upsert_personas_batch(pers)
            conn.commit()

            if resultados.get('integrantes'):
                _, grupos_b = resultados['integrantes']
                if grupos_b:
                    conn.executemany(
                        'INSERT OR IGNORE INTO grupos (cedula,grupo,facultad,tipo_miembro) VALUES(?,?,?,?)',
                        grupos_b,
                    )
                    self.progreso.emit(f"Insertados {len(grupos_b)} integrantes de grupos")

            if resultados.get('extensiones'):
                _, ext_b = resultados['extensiones']
                if ext_b:
                    conn.executemany(
                        '''INSERT OR REPLACE INTO extensiones
                           (cedula,actividad,tipo,modalidad,estado,fecha_inicio,fecha_fin,
                            año,poblacion,grupo,facultad,financiacion_interna,financiacion_externa,fuente,
                            datos_adicionales)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                        ext_b,
                    )
                    self.progreso.emit(f"Insertadas/actualizadas {len(ext_b)} extensiones")

            if resultados.get('produccion'):
                _, pub_b = resultados['produccion']
                if pub_b:
                    conn.executemany(
                        '''INSERT OR REPLACE INTO publicaciones
                           (cedula,titulo,revista_libro,doi_url,issn_isbn,año,
                            tipo,categoria,estado,grupo,fuente,datos_adicionales)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''',
                        pub_b,
                    )
                    self.progreso.emit(f"Insertadas/actualizadas {len(pub_b)} publicaciones")

            if resultados.get('trabajos_grado'):
                _, tg_b = resultados['trabajos_grado']
                if tg_b:
                    conn.executemany(
                        '''INSERT OR REPLACE INTO trabajos_grado
                           (cedula_director,nombre_director,cedula_estudiante,nombre_estudiante,
                            titulo,programa,año,calificacion,fecha_sustentacion,facultad,fuente)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?)''',
                        tg_b,
                    )
                    self.progreso.emit(f"Insertados/actualizados {len(tg_b)} trabajos de grado")

            if resultados.get('libros'):
                _, libros_b = resultados['libros']
                if libros_b:
                    conn.executemany(
                        '''INSERT OR REPLACE INTO publicaciones
                           (cedula,titulo,revista_libro,doi_url,issn_isbn,año,
                            tipo,categoria,estado,grupo,fuente)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?)''',
                        libros_b,
                    )
                    self.progreso.emit(f"Insertados/actualizados {len(libros_b)} libros/capítulos")

            if resultados.get('productos_innovacion'):
                _, pi_b = resultados['productos_innovacion']
                if pi_b:
                    conn.executemany(
                        '''INSERT OR REPLACE INTO productos_innovacion
                           (cedula,tipo_producto,nombre,descripcion,año,estado,grupo,fuente,datos_adicionales)
                           VALUES(?,?,?,?,?,?,?,?,?)''',
                        pi_b,
                    )
                    self.progreso.emit(f"Insertados/actualizados {len(pi_b)} productos de innovación")

            conn.commit()

            # ── Cargas secuenciales (dependen del estado de la BD) ──────────
            self.cargar_proyectos()
            self.cargar_propiedad_intelectual()

            self.progreso.emit("Detectando y consolidando duplicados…")
            consolidaciones = self.db.consolidar_duplicados()
            if consolidaciones:
                self.progreso.emit(
                    f"Consolidados {len(consolidaciones)} grupos de duplicados"
                )
                self.duplicados_consolidados.emit(consolidaciones)
            else:
                self.progreso.emit("No se encontraron duplicados")

            stats = self.db.obtener_estadisticas()
            self.db.guardar_sello_carga(archivos_fuente)
            self.finalizado.emit(stats)
        except Exception as e:
            import traceback
            self.progreso.emit(f"Error en carga: {e}")
            self.progreso.emit(traceback.format_exc()[:600])
        finally:
            conn.execute("PRAGMA synchronous=NORMAL")
            conn.commit()

    def _norm_key(k: str) -> str:
        """Normaliza una clave de columna para búsqueda robusta."""
        return normalizar_columna(k)

    @classmethod
    def _col(cls, row: dict, *keys) -> str:
        """
        Primer valor no-vacío buscando la clave tanto en su forma original
        como normalizada (sin tildes, espacios→_, minúsculas).
        Funciona con dicts de to_dict('records') Y con _asdict() de itertuples.
        """
        for k in keys:
            v = row.get(k)
            if v is not None and str(v).strip() not in ('', 'nan', 'None'):
                return v
            k_norm = cls._norm_key(k)
            v = row.get(k_norm)
            if v is not None and str(v).strip() not in ('', 'nan', 'None'):
                return v
        return ''

    @classmethod
    def _normalizar_cols(cls, df):
        """Devuelve el mismo DataFrame con columnas normalizadas en sus nombres."""
        df = df.copy()
        df.columns = [cls._norm_key(c) for c in df.columns]
        return df

    @staticmethod
    def _columnas_extra_json(df, columnas_conocidas: set) -> pd.Series:
        """
        Serializa por fila las columnas del Excel que no están en
        `columnas_conocidas` (columnas nuevas/no mapeadas), para no perder esa
        información aunque todavía no tenga una columna propia en la BD.
        Devuelve '' por fila cuando no hay columnas extra o todas vienen vacías.
        """
        extra_cols = [c for c in df.columns if c not in columnas_conocidas]
        if not extra_cols:
            return pd.Series([''] * len(df), index=df.index)

        def _serializar(fila):
            datos = {c: fila[c] for c in extra_cols if limpiar_texto(fila[c])}
            return json.dumps(datos, ensure_ascii=False) if datos else ''

        return df[extra_cols].apply(_serializar, axis=1)

    @staticmethod
    def _anio(valor):
        """Extrae año entero de un valor; devuelve None si no es posible."""
        if not valor or (isinstance(valor, float) and valor != valor):
            return None
        try:
            return int(str(valor).strip()[:4])
        except (ValueError, TypeError):
            return None

    # ── Integrantes ────────────────────────────────────────────────────────
    def _extraer_integrantes(self):
        archivos = [
            'Listado Integrantes Grupos de Investigación UTP  080825.xlsx',
            'Listado Integrantes Grupos de Investigación UTP - 080825.xlsx',
            'Listado Integrantes Grupos de Investigacion UTP  080825.xlsx',
            'Listado Integrantes Grupos de Investigacion UTP - 080825.xlsx',
            'integrantes.xlsx',
        ]
        for archivo in archivos:
            ruta = self.archivos_directorio / archivo
            if not ruta.exists():
                continue
            df = self._normalizar_cols(
                pd.read_excel(ruta, engine='openpyxl', dtype=str).fillna('')
            )
            cedulas    = df.get('numero_documento', df.get('cedula', pd.Series(dtype=str))).apply(limpiar_cedula)
            nombres    = df.get('nombres', df.get('nombre', pd.Series(dtype=str))).apply(limpiar_texto)
            grupos_s   = df.get('nombre_grupo', df.get('grupo', pd.Series(dtype=str))).apply(limpiar_texto)
            facultades = df.get('facultad', pd.Series(dtype=str)).apply(limpiar_texto)
            emails     = df.get('email', pd.Series(dtype=str)).apply(limpiar_texto)
            tipos      = df.get('tipo', pd.Series(dtype=str)).apply(limpiar_texto)
            mask = (cedulas.str.len() > 0) & (nombres.str.len() > 0)
            personas_b = list(zip(cedulas[mask], nombres[mask], emails[mask], facultades[mask], tipos[mask]))
            grupos_b   = [
                (c, g, f, t)
                for c, g, f, t in zip(cedulas[mask], grupos_s[mask], facultades[mask], tipos[mask])
                if g
            ]
            return personas_b, grupos_b
        return [], []

    def cargar_integrantes(self):
        self.progreso.emit('Cargando integrantes de grupos...')
        personas_b, grupos_b = self._extraer_integrantes()
        if personas_b:
            conn = self.db.conn
            self.db._upsert_personas_batch(personas_b)
            conn.executemany(
                'INSERT OR IGNORE INTO grupos (cedula,grupo,facultad,tipo_miembro) VALUES(?,?,?,?)',
                grupos_b,
            )
            conn.commit()
            self.progreso.emit(f'Cargados {len(personas_b)} integrantes de grupos')

    # ── Extensiones ──────────────────────────────────────────────
    def _extraer_extensiones(self):
        archivos = [
            'Consolidado Extensión 2024.xlsx',
            'Actividades Extensión enerojulio.xlsx',
            'Actividades Extensión (enero-julio).xlsx',
        ]
        personas_b, ext_b = [], []
        for archivo in archivos:
            ruta = self.archivos_directorio / archivo
            if not ruta.exists():
                continue
            df = self._normalizar_cols(
                pd.read_excel(ruta, sheet_name='Consolidado', engine='openpyxl', dtype=str).fillna('')
            )

            cedulas = df.get('cedula', pd.Series(dtype=str)).apply(limpiar_cedula)
            mask = cedulas.str.len() > 0
            if not mask.any():
                continue

            dv = df[mask]
            cv = cedulas[mask]
            nombres   = dv.get('nombre_responsable', pd.Series(dtype=str)).apply(limpiar_texto).replace('', 'Sin nombre')
            facultades = (dv.get('facultad_dependencia') if 'facultad_dependencia' in dv.columns
                          else dv.get('facultad', pd.Series(dtype=str))).apply(limpiar_texto)
            fis       = dv.get('fecha_inicial', pd.Series(dtype=str)).astype(str)
            anios     = fis.str[:4].apply(self._anio)

            grupos_s  = (dv.get('grupo_semillero_de_investigacion') if 'grupo_semillero_de_investigacion' in dv.columns
                         else dv.get('grupo', pd.Series(dtype=str))).apply(limpiar_texto)
            extra     = self._columnas_extra_json(dv, COLUMNAS_CONOCIDAS_POR_CATEGORIA['extension'])

            personas_b.extend(zip(cv, nombres, pd.Series('', index=dv.index), facultades, pd.Series('Responsable', index=dv.index)))
            ext_b.extend(zip(
                cv,
                dv.get('nombre_actividad', pd.Series(dtype=str)).apply(limpiar_texto),
                dv.get('tipo', pd.Series(dtype=str)).apply(limpiar_texto),
                dv.get('modalidad', pd.Series(dtype=str)).apply(limpiar_texto),
                dv.get('estado', pd.Series(dtype=str)).apply(limpiar_texto),
                fis,
                dv.get('fecha_final', pd.Series(dtype=str)).apply(limpiar_texto),
                anios,
                dv.get('poblacion_beneficiaria', pd.Series(dtype=str)).apply(limpiar_texto),
                grupos_s,
                facultades,
                dv.get('financiacion_interna', pd.Series(dtype=str)).apply(limpiar_texto),
                dv.get('fuente_financiacion_externa', pd.Series(dtype=str)).apply(limpiar_texto),
                [archivo] * len(dv),
                extra,
            ))
        return personas_b, ext_b

    def cargar_extensiones(self):
        self.progreso.emit('Cargando extensiones...')
        personas_b, ext_b = self._extraer_extensiones()
        if ext_b:
            conn = self.db.conn
            self.db._upsert_personas_batch(personas_b)
            conn.executemany(
                '''INSERT OR REPLACE INTO extensiones
                   (cedula,actividad,tipo,modalidad,estado,fecha_inicio,fecha_fin,
                    año,poblacion,grupo,facultad,financiacion_interna,financiacion_externa,fuente,
                    datos_adicionales)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                ext_b,
            )
            conn.commit()
            self.progreso.emit(f'Cargadas {len(ext_b)} extensiones')

    # ── Producción / Publicaciones ──────────────────────────────
    def _extraer_produccion(self):
        archivos = [
            'BASE DATOS PRODUCCIÓN 2024.xlsx',
            'BASE DATOS PRODUCCIÓN  2025  CIARP.xlsx',
            'BASE DATOS PRODUCCIÓN  2025 - CIARP.xlsx',
        ]
        personas_b, pub_b = [], []
        for archivo in archivos:
            ruta = self.archivos_directorio / archivo
            if not ruta.exists():
                continue
            xls = pd.ExcelFile(ruta, engine='openpyxl')
            for sheet in xls.sheet_names:
                df = self._normalizar_cols(
                    pd.read_excel(xls, sheet_name=sheet, dtype=str).fillna('')
                )

                cedulas = df.get('cedula', pd.Series(dtype=str)).apply(limpiar_cedula)
                mask = cedulas.str.len() > 0
                if not mask.any():
                    continue

                dv = df[mask]
                cv = cedulas[mask]
                nombres   = (dv.get('autores') if 'autores' in dv.columns
                             else dv.get('autor') if 'autor' in dv.columns
                             else dv.get('nombre', pd.Series(dtype=str))).apply(limpiar_texto)
                facultades = (dv.get('dependencia') if 'dependencia' in dv.columns
                              else dv.get('facultad', pd.Series(dtype=str))).apply(limpiar_texto)
                extra = self._columnas_extra_json(dv, COLUMNAS_CONOCIDAS_POR_CATEGORIA['produccion_2024'])

                fuente = f'{archivo} :: {sheet}'
                personas_b.extend(zip(cv, nombres, pd.Series('', index=dv.index), facultades, pd.Series('Autor', index=dv.index)))
                pub_b.extend(zip(
                    cv,
                    (dv.get('nombre_del_trabajo') if 'nombre_del_trabajo' in dv.columns
                     else dv.get('titulo', pd.Series(dtype=str))).apply(limpiar_texto),
                    (dv.get('revista_o_libro') if 'revista_o_libro' in dv.columns
                     else dv.get('revista_libro', pd.Series(dtype=str))).apply(limpiar_texto),
                    (dv.get('doi_url') if 'doi_url' in dv.columns
                     else dv.get('doi', pd.Series(dtype=str))).apply(limpiar_texto),
                    (dv.get('issn_isbn') if 'issn_isbn' in dv.columns
                     else dv.get('issn', pd.Series(dtype=str))).apply(limpiar_texto),
                    (dv.get('ano_de_la_publicacion') if 'ano_de_la_publicacion' in dv.columns
                     else dv.get('ano', pd.Series(dtype=str))).apply(self._anio),
                    dv.get('tipo', pd.Series(dtype=str)).apply(limpiar_texto),
                    dv.get('categoria', pd.Series(dtype=str)).apply(limpiar_texto),
                    dv.get('estado', pd.Series(dtype=str)).apply(limpiar_texto),
                    dv.get('grupo', pd.Series(dtype=str)).apply(limpiar_texto),
                    [fuente] * len(dv),
                    extra,
                ))
        return personas_b, pub_b

    def cargar_produccion(self):
        self.progreso.emit('Cargando publicaciones...')
        personas_b, pub_b = self._extraer_produccion()
        if pub_b:
            conn = self.db.conn
            self.db._upsert_personas_batch(personas_b)
            conn.executemany(
                '''INSERT OR REPLACE INTO publicaciones
                   (cedula,titulo,revista_libro,doi_url,issn_isbn,año,
                    tipo,categoria,estado,grupo,fuente,datos_adicionales)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?)''',
                pub_b,
            )
            conn.commit()
            self.progreso.emit(f'Cargadas {len(pub_b)} publicaciones')

    # ── Trabajos de grado ─────────────────────────────────────────
    def _extraer_trabajos_grado(self):
        archivos = [
            'Trabajos Grado  Trabajo de Grado.xlsx',
            'TrabajosGrado_TrabajoDeGrado 2024.xlsx',
            'Trabajos Grado - Trabajo de Grado.xlsx',
        ]
        personas_b, tg_b = [], []
        for archivo in archivos:
            ruta = self.archivos_directorio / archivo
            if not ruta.exists():
                continue
            df = pd.read_excel(ruta, engine='openpyxl', header=None)
            director_actual = cedula_director = None
            for row in df.itertuples(index=False):
                col_a = limpiar_texto(str(row[0]) if pd.notna(row[0]) else '')
                if col_a and 'Nombre del trabajo' not in col_a:
                    ultima = row[-1] if len(row) > 1 else None
                    if pd.notna(ultima):
                        ced = limpiar_cedula(str(ultima))
                        if ced and len(ced) >= 6:
                            director_actual = col_a
                            cedula_director = ced
                            personas_b.append((ced, col_a, '', '', 'Director'))
                            continue
                if director_actual and cedula_director:
                    titulo = limpiar_texto(str(row[0]) if pd.notna(row[0]) else '')
                    if not titulo or titulo == 'Nombre del trabajo de grado':
                        continue
                    fecha = limpiar_texto(str(row[5]) if len(row) > 5 and pd.notna(row[5]) else '')
                    anio = None
                    if fecha:
                        try:
                            anio = int(fecha.split('/')[-1]) if '/' in fecha else int(fecha.split('-')[0])
                        except (ValueError, IndexError):
                            pass
                    # Columnas del Excel: A=Título, B=Documento Estudiante,
                    # C=Estudiante (nombre), D=Programa, E=Nota, F=Fecha
                    # Sustentación, G=Facultad.
                    tg_b.append((
                        cedula_director, director_actual,
                        limpiar_cedula(str(row[1]) if len(row) > 1 and pd.notna(row[1]) else ''),
                        limpiar_texto(str(row[2]) if len(row) > 2 and pd.notna(row[2]) else ''),
                        titulo,
                        limpiar_texto(str(row[3]) if len(row) > 3 and pd.notna(row[3]) else ''),
                        anio,
                        limpiar_texto(str(row[4]) if len(row) > 4 and pd.notna(row[4]) else ''),
                        fecha,
                        limpiar_texto(str(row[6]) if len(row) > 6 and pd.notna(row[6]) else ''),
                        archivo,
                    ))
        return personas_b, tg_b

    def cargar_trabajos_grado(self):
        self.progreso.emit('Cargando trabajos de grado...')
        personas_b, tg_b = self._extraer_trabajos_grado()
        if tg_b:
            conn = self.db.conn
            self.db._upsert_personas_batch(personas_b)
            conn.executemany(
                '''INSERT OR REPLACE INTO trabajos_grado
                   (cedula_director,nombre_director,cedula_estudiante,nombre_estudiante,
                    titulo,programa,año,calificacion,fecha_sustentacion,facultad,fuente)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?)''',
                tg_b,
            )
            conn.commit()
            self.progreso.emit(f'Cargados {len(tg_b)} trabajos de grado')

    # ── Libros ────────────────────────────────────────────────────
    def _extraer_libros(self):
        archivos = [
            'Reporte de libros y capítulos publicados.xlsx',
            'Reporte_libros.xlsx',
        ]
        personas_b, pub_b = [], []
        for archivo in archivos:
            ruta = self.archivos_directorio / archivo
            if not ruta.exists():
                continue
            df = pd.read_excel(ruta, engine='openpyxl', header=None)
            titulo_actual = tipo_actual = None
            autores_libro = []

            def _flush_libro():
                for aut in autores_libro:
                    pub_b.append((
                        aut['cedula'], titulo_actual, '', '', '', None,
                        'LIBRO', tipo_actual or 'LIBRO', '', '', archivo,
                    ))

            for row in df.itertuples(index=False):
                nombre = limpiar_texto(str(row[0]) if pd.notna(row[0]) else '')
                titulo = limpiar_texto(str(row[1]) if len(row) > 1 and pd.notna(row[1]) else '')
                tipo   = limpiar_texto(str(row[3]) if len(row) > 3 and pd.notna(row[3]) else '')
                cedula = limpiar_cedula(str(row[4]) if len(row) > 4 and pd.notna(row[4]) else '')
                if titulo and titulo != titulo_actual:
                    _flush_libro()
                    titulo_actual = titulo
                    tipo_actual   = tipo if tipo else 'LIBRO'
                    autores_libro = []
                if nombre and cedula:
                    personas_b.append((cedula, nombre, '', '', 'Autor'))
                    autores_libro.append({'cedula': cedula})
            _flush_libro()
        return personas_b, pub_b

    def cargar_libros(self):
        self.progreso.emit('Cargando libros y capítulos...')
        personas_b, pub_b = self._extraer_libros()
        if pub_b:
            conn = self.db.conn
            self.db._upsert_personas_batch(personas_b)
            conn.executemany(
                '''INSERT OR REPLACE INTO publicaciones
                   (cedula,titulo,revista_libro,doi_url,issn_isbn,año,
                    tipo,categoria,estado,grupo,fuente)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?)''',
                pub_b,
            )
            conn.commit()
            self.progreso.emit(f'Cargados {len(pub_b)} libros/capítulos')

    # ── Productos de innovación ───────────────────────────────────────
    def _extraer_productos_innovacion(self):
        ruta = self.archivos_directorio / 'info_productos_innovacion.xlsx'
        if not ruta.exists():
            return [], []
        df = self._normalizar_cols(
            pd.read_excel(ruta, engine='openpyxl', dtype=str).fillna('')
        )
        nombres = (df.get('nombre') if 'nombre' in df.columns
                   else df.get('titulo') if 'titulo' in df.columns
                   else df.get('producto', pd.Series(dtype=str))).apply(limpiar_texto)
        mask = nombres.str.len() > 0
        if not mask.any():
            return [], []
        dv = df[mask]
        nv = nombres[mask]
        anios = (dv.get('ano_de_registro') if 'ano_de_registro' in dv.columns
                 else dv.get('ano', pd.Series(dtype=str))).apply(self._anio)
        extra = self._columnas_extra_json(dv, COLUMNAS_CONOCIDAS_POR_CATEGORIA['innovacion'])
        batch = list(zip(
            ['0000000'] * len(dv),
            (dv.get('tipo_de_producto') if 'tipo_de_producto' in dv.columns
             else dv.get('tipo', pd.Series(dtype=str))).apply(limpiar_texto),
            nv,
            dv.get('descripcion', pd.Series(dtype=str)).apply(limpiar_texto),
            anios,
            dv.get('estado', pd.Series(dtype=str)).apply(limpiar_texto),
            (dv.get('grupo_de_investigacion') if 'grupo_de_investigacion' in dv.columns
             else dv.get('grupo', pd.Series(dtype=str))).apply(limpiar_texto),
            ['info_productos_innovacion.xlsx'] * len(dv),
            extra,
        ))
        return [], batch

    def cargar_productos_innovacion(self):
        self.progreso.emit('Cargando productos de innovación...')
        _, batch = self._extraer_productos_innovacion()
        if batch:
            conn = self.db.conn
            conn.executemany(
                '''INSERT OR REPLACE INTO productos_innovacion
                   (cedula,tipo_producto,nombre,descripcion,año,estado,grupo,fuente,datos_adicionales)
                   VALUES(?,?,?,?,?,?,?,?,?)''',
                batch,
            )
            conn.commit()
            self.progreso.emit(f'Cargados {len(batch)} productos de innovación')

    def cargar_proyectos(self):
        self.progreso.emit("Cargando proyectos de investigación...")
        
        # ARCHIVOS POSIBLES - buscar con variaciones
        archivos_posibles = [
            "Proyectos de investigación registrados en 2024.xlsx",
            "Proyectos de investigacion registrados en 2024.xlsx",
            "proyectos_investigacion_2024.xlsx",
            "proyectos 2024.xlsx"
        ]
        
        archivo_encontrado = None
        for archivo in archivos_posibles:
            ruta = self.archivos_directorio / archivo
            if ruta.exists():
                archivo_encontrado = archivo
                break
        
        if not archivo_encontrado:
            self.progreso.emit("⚠ No se encontró archivo de proyectos de investigación")
            return
        
        ruta = self.archivos_directorio / archivo_encontrado
        
        try:
            # USAR DETECCIÓN AUTOMÁTICA DE ENCABEZADOS (como investigacion.py)
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", UserWarning)
                df_raw = pd.read_excel(ruta, sheet_name=0, header=None, dtype=str, engine="openpyxl")
            
            df_raw = df_raw.fillna("")
            
            # Detectar fila de encabezados
            hdr_row = self._find_header_row_proyectos(df_raw)
            
            if hdr_row is None:
                self.progreso.emit(f"⚠ No se detectó encabezado en {archivo_encontrado}. Intentando lectura estándar...")
                # Intento con lectura normal
                df = pd.read_excel(ruta, engine='openpyxl')
            else:
                self.progreso.emit(f"✓ Encabezado detectado en fila {hdr_row+1}")
                
                # Mapear columnas
                colmap = self._pick_columns_proyectos(df_raw.iloc[hdr_row, :])
                self.progreso.emit(f"Columnas detectadas: {sum(1 for v in colmap.values() if v is not None)}/{len(colmap)}")
                
                # Construir DataFrame
                df = pd.DataFrame()
                for col_name, col_idx in colmap.items():
                    if col_idx is not None:
                        df[col_name] = df_raw.iloc[hdr_row+1:, col_idx].astype(str).str.strip()
            
            count = 0
            for _, row in df.iterrows():
                # Obtener campos (ahora con nombres normalizados si usó detección)
                if 'RESPONSABLE' in df.columns:
                    responsable = limpiar_texto(row.get('RESPONSABLE', ''))
                    cedula = limpiar_cedula(row.get('CEDULA', ''))
                    titulo = limpiar_texto(row.get('TITULO', ''))
                    objetivo = limpiar_texto(row.get('OBJETIVO', ''))
                    codigo_cie = limpiar_texto(row.get('CODIGO_CIE', ''))
                    tipo = limpiar_texto(row.get('TIPO_INV', ''))
                    año = row.get('ANIO', '')
                    fecha_inicio = limpiar_texto(row.get('FECHA_INICIO', ''))
                    fecha_fin = limpiar_texto(row.get('FECHA_FINAL', ''))
                    estado = limpiar_texto(row.get('ESTADO', ''))
                    facultad = limpiar_texto(row.get('FACULTAD', ''))
                    grupo = limpiar_texto(row.get('GRUPO', ''))
                    valor = limpiar_texto(row.get('VALOR_APROBADO', ''))
                else:
                    # Fallback: buscar por nombres de columnas originales
                    cedula = limpiar_cedula(
                        row.get('Cedula') or row.get('Cédula') or row.get('CEDULA') or 
                        row.get('Documento') or row.get('Número de documento') or ''
                    )
                    responsable = limpiar_texto(
                        row.get('Responsable') or row.get('RESPONSABLE') or 
                        row.get('Investigador principal') or row.get('Nombre') or ''
                    )
                    titulo = limpiar_texto(
                        row.get('Título') or row.get('Titulo') or row.get('TITULO') or
                        row.get('Nombre proyecto') or ''
                    )
                    objetivo = limpiar_texto(row.get('Objetivo') or row.get('OBJETIVO') or '')
                    codigo_cie = limpiar_texto(row.get('Código CIE') or row.get('Codigo CIE') or '')
                    tipo = limpiar_texto(row.get('Tipo') or row.get('TIPO') or '')
                    año = row.get('Año') or row.get('AÑO') or row.get('Ano')
                    fecha_inicio = limpiar_texto(row.get('Fecha inicio') or row.get('FECHA INICIO') or '')
                    fecha_fin = limpiar_texto(row.get('Fecha final') or row.get('FECHA FINAL') or '')
                    estado = limpiar_texto(row.get('Estado') or row.get('ESTADO') or '')
                    facultad = limpiar_texto(row.get('Facultad') or row.get('FACULTAD') or '')
                    grupo = limpiar_texto(row.get('Grupo') or row.get('GRUPO') or '')
                    valor = limpiar_texto(row.get('Valor aprobado') or row.get('VALOR APROBADO') or '')
                
                if not titulo:
                    continue
                
                if not cedula:
                    cedula = '0000000'
                
                if responsable:
                    self.db.insertar_persona(cedula, responsable, '', facultad, 'Investigador')
                
                # Convertir año
                if pd.notna(año) and año:
                    try:
                        año = int(str(año).split('.')[0])
                    except:
                        año = None
                
                cedula_principal = self.db.obtener_cedula_principal(cedula)
                cursor = self.db.conn.cursor()
                cursor.execute('''
                    INSERT OR REPLACE INTO proyectos
                    (cedula, responsable, titulo, objetivo, codigo_cie, tipo, año,
                     fecha_inicio, fecha_fin, estado, facultad, grupo, valor_aprobado, fuente)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    cedula_principal, responsable, titulo, objetivo, codigo_cie, 
                    tipo, año, fecha_inicio, fecha_fin, estado, facultad, 
                    grupo, valor, archivo_encontrado
                ))
                count += 1
            
            self.db.conn.commit()
            self.progreso.emit(f"✓ Cargados {count} proyectos de investigación desde {archivo_encontrado}")
        except Exception as e:
            self.progreso.emit(f"Error en proyectos: {str(e)}")
    
    def _find_header_row_proyectos(self, df: pd.DataFrame, max_scan: int = 50) -> int:
        """Detecta la fila de encabezados en proyectos"""
        keys_any = {"responsable", "responsables", "investigador principal"}
        keys_id = {"cedula", "documento"}
        other_ok = {"codigo cie", "objetivo", "titulo"}
        
        rows = min(max_scan, df.shape[0])
        for r in range(rows):
            vals = [norm_text(str(v)) for v in df.iloc[r, :].tolist()]
            have_resp = any(any(k in v for k in keys_any) for v in vals)
            have_id = any(any(k in v for k in keys_id) for v in vals)
            have_other = any(any(k in v for k in other_ok) for v in vals)
            if have_resp and (have_id or have_other):
                return r
        return None
    
    def _pick_columns_proyectos(self, header: pd.Series) -> dict:
        """Mapea columnas de proyectos"""
        cols = {
            "RESPONSABLE": None, "CEDULA": None, "TITULO": None, "OBJETIVO": None,
            "CODIGO_CIE": None, "ANIO": None, "FECHA_INICIO": None, "FECHA_FINAL": None,
            "ESTADO": None, "TIPO_INV": None, "VALOR_APROBADO": None, 
            "FACULTAD": None, "GRUPO": None
        }
        
        for j, val in header.items():
            key = norm_text(str(val))
            if cols["RESPONSABLE"] is None and ("responsable" in key or "investigador principal" in key):
                cols["RESPONSABLE"] = j
            elif cols["CEDULA"] is None and ("cedula" in key or "documento" in key):
                cols["CEDULA"] = j
            elif cols["TITULO"] is None and ("titulo" in key or "nombre del proyecto" in key or "proyecto" in key):
                cols["TITULO"] = j
            elif cols["OBJETIVO"] is None and "objetivo" in key:
                cols["OBJETIVO"] = j
            elif cols["CODIGO_CIE"] is None and "cie" in key:
                cols["CODIGO_CIE"] = j
            elif cols["ANIO"] is None and ("ano" in key or "year" in key):
                cols["ANIO"] = j
            elif cols["FECHA_INICIO"] is None and "inicio" in key:
                cols["FECHA_INICIO"] = j
            elif cols["FECHA_FINAL"] is None and ("final" in key or "finalizacion" in key):
                cols["FECHA_FINAL"] = j
            elif cols["ESTADO"] is None and "estado" in key:
                cols["ESTADO"] = j
            elif cols["TIPO_INV"] is None and "tipo" in key:
                cols["TIPO_INV"] = j
            elif cols["VALOR_APROBADO"] is None and ("valor" in key or "aprobado" in key):
                cols["VALOR_APROBADO"] = j
            elif cols["FACULTAD"] is None and "facultad" in key:
                cols["FACULTAD"] = j
            elif cols["GRUPO"] is None and "grupo" in key:
                cols["GRUPO"] = j
        
        return cols
    
    def cargar_propiedad_intelectual(self):
        self.progreso.emit("Cargando registros de propiedad intelectual...")
        archivos_posibles = [
            "CGT0104  No de productos resultados de investigacion 31072025.xlsx",
            "CGT0104  No de productos resultados de investigacion 31122024.xlsx"
        ]
        
        count = 0
        for archivo in archivos_posibles:
            ruta = self.archivos_directorio / archivo
            if not ruta.exists():
                continue
            
            try:
                wb = pd.ExcelFile(ruta, engine='openpyxl')
                for sheet_name in wb.sheet_names:
                    if 'Soporte' in sheet_name or 'Reg Propiedad' in sheet_name:
                        df = pd.read_excel(ruta, sheet_name=sheet_name, engine='openpyxl')
                        
                        columnas_conocidas = COLUMNAS_CONOCIDAS_POR_CATEGORIA['cgt0104_2025']
                        col_extra_por_norm = {
                            c: c for c in df.columns
                            if normalizar_columna(c) not in columnas_conocidas
                        }

                        for _, row in df.iterrows():
                            responsable = limpiar_texto(row.get('Responsables') or row.get('Responsable') or '')
                            tipo_producto = limpiar_texto(row.get('Tipo de producto') or '')
                            nombre = limpiar_texto(row.get('Nombre del producto') or row.get('Nombre') or '')

                            if not responsable and not nombre:
                                continue

                            datos_extra = {
                                c: limpiar_texto(row.get(c))
                                for c in col_extra_por_norm
                                if limpiar_texto(row.get(c))
                            }
                            extra_json = json.dumps(datos_extra, ensure_ascii=False) if datos_extra else ''

                            cursor = self.db.conn.cursor()
                            cursor.execute('''
                                INSERT OR REPLACE INTO propiedad_intelectual
                                (responsable, tipo_producto, tipo_patente, nombre_producto,
                                 numero_registro, proyecto, fecha_aprobacion, entidad, facultad, fuente,
                                 datos_adicionales)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            ''', (
                                responsable,
                                tipo_producto,
                                limpiar_texto(row.get('Tipo de patente') or ''),
                                nombre,
                                limpiar_texto(row.get('No de registro') or ''),
                                limpiar_texto(row.get('Proyecto de investigación') or ''),
                                limpiar_texto(row.get('Fecha de aprobación') or ''),
                                limpiar_texto(row.get('Entidad que lo expide') or ''),
                                limpiar_texto(row.get('Facultad') or ''),
                                f"{archivo} :: {sheet_name}",
                                extra_json,
                            ))
                            count += 1
                
                self.db.conn.commit()
            except Exception as e:
                self.progreso.emit(f"Error en {archivo}: {str(e)}")
        
        if count > 0:
            self.progreso.emit(f"Cargados {count} registros de propiedad intelectual")

# ==================== INTERFAZ ====================

class VistaBusqueda(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self._analisis_gruplac = AnalisisDuplicados(db.conn)
        self._cache_gruplac = None
        self.setup_ui()

    def _obtener_cache_gruplac(self):
        """Carga (una sola vez por sesión) los integrantes de todos los grupos
        GrupLAC desde reports/excel, para verificar coherencia interno↔GrupLAC."""
        if self._cache_gruplac is None:
            self._cache_gruplac = self._analisis_gruplac._cargar_integrantes_gruplac()
        return self._cache_gruplac
    
    def setup_ui(self):
        layout = QHBoxLayout()
        
        panel_izq = QWidget()
        layout_izq = QVBoxLayout(panel_izq)
        
        self.input_busqueda = QLineEdit()
        self.input_busqueda.setPlaceholderText("Buscar por nombre o cédula...")
        self.input_busqueda.returnPressed.connect(self.buscar)
        
        btn_buscar = QPushButton("Buscar")
        btn_buscar.clicked.connect(self.buscar)
        
        btn_consolidar = QPushButton("Consolidar Duplicados")
        btn_consolidar.clicked.connect(self.consolidar_duplicados_manual)
        btn_consolidar.setStyleSheet("background-color: #e67e22; color: white;")
        
        self.tabla_resultados = QTableWidget()
        self.tabla_resultados.setColumnCount(5)
        self.tabla_resultados.setHorizontalHeaderLabels(['Nombre', 'Cédula', 'Email', 'Facultad', 'Tipo'])
        self.tabla_resultados.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tabla_resultados.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tabla_resultados.horizontalHeader().setStretchLastSection(True)
        self.tabla_resultados.doubleClicked.connect(self.mostrar_detalle)
        
        layout_izq.addWidget(QLabel("Búsqueda de Personas"))
        layout_izq.addWidget(self.input_busqueda)
        layout_izq.addWidget(btn_buscar)
        layout_izq.addWidget(btn_consolidar)
        layout_izq.addWidget(self.tabla_resultados)
        
        panel_der = QWidget()
        layout_der = QVBoxLayout(panel_der)
        
        self.texto_detalle = QTextEdit()
        self.texto_detalle.setReadOnly(True)
        
        layout_der.addWidget(QLabel("Detalle de Persona"))
        layout_der.addWidget(self.texto_detalle)
        
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(panel_izq)
        splitter.addWidget(panel_der)
        splitter.setSizes([400, 600])
        
        layout.addWidget(splitter)
        self.setLayout(layout)
    
    def consolidar_duplicados_manual(self):
        consolidaciones = self.db.consolidar_duplicados()
        
        if consolidaciones:
            mensaje = f"✓ Se consolidaron {len(consolidaciones)} persona(s):\n\n"
            for grupo in consolidaciones:
                mensaje += f"- {grupo['nombre']}\n"
                mensaje += f"  Cédula principal: {grupo['cedula_principal']}\n"
                mensaje += f"  Cédulas consolidadas: {', '.join(grupo['cedulas'])}\n\n"
            
            QMessageBox.information(self, "Consolidación Exitosa", mensaje)
            
            if self.input_busqueda.text():
                self.buscar()
        else:
            QMessageBox.information(self, "Sin Duplicados", "No se encontraron duplicados para consolidar.")
    
    def buscar(self):
        termino = self.input_busqueda.text().strip()
        if not termino:
            return
        
        resultados = self.db.buscar_personas(termino)
        self.tabla_resultados.setRowCount(len(resultados))
        
        for row, persona in enumerate(resultados):
            self.tabla_resultados.setItem(row, 0, QTableWidgetItem(persona[1] or ''))
            self.tabla_resultados.setItem(row, 1, QTableWidgetItem(persona[0] or ''))
            self.tabla_resultados.setItem(row, 2, QTableWidgetItem(persona[2] or ''))
            self.tabla_resultados.setItem(row, 3, QTableWidgetItem(persona[3] or ''))
            self.tabla_resultados.setItem(row, 4, QTableWidgetItem(persona[4] or ''))
    
    def mostrar_detalle(self, index):
        row = index.row()
        cedula = self.tabla_resultados.item(row, 1).text()
        detalle = self.db.obtener_detalle_persona(cedula)
        if not detalle or "info" not in detalle:
            return

        info = detalle["info"]

        def _s(v):
            return str(v) if v else "—"

        def _extra_html(fila):
            """Si la fila trae 'datos_adicionales' (columnas nuevas guardadas
            como JSON, ver views/vista_inicio.py), las muestra como nota."""
            crudo = fila[-1] if fila else None
            if not crudo:
                return ""
            try:
                datos = json.loads(crudo)
            except (TypeError, ValueError):
                return ""
            if not datos:
                return ""
            partes = "; ".join(f"{k}: {v}" for k, v in datos.items())
            return f"<br><span style='color:#888;font-size:9px;'>+ {partes}</span>"

        def _seccion(titulo, color, items_html):
            return (
                f"<div style='margin-top:10px;'>"
                f"<div style='background:{color};color:white;padding:4px 8px;"
                f"border-radius:4px;font-weight:bold;font-size:12px;'>{titulo}</div>"
                f"<div style='padding:4px 8px;'>{items_html}</div></div>"
            )

        html = (
            "<html><body style='font-family:Arial,sans-serif;font-size:11px;"
            "max-width:100%;word-wrap:break-word;'>"
            "<div style='background:#1a365d;color:white;padding:8px;border-radius:6px;"
            "margin-bottom:10px;'>"
            f"<b style='font-size:13px;'>{_s(info['nombre'])}</b><br>"
            f"<span style='font-size:10px;'>Cédula: {_s(info['cedula'])} | "
            f"Facultad: {_s(info['facultad'])}<br>"
            f"Email: {_s(info['email'])} | Tipo: {_s(info['tipo'])}</span></div>"
        )

        grupos_todos = detalle.get("grupos") or []
        grupos_invest = [g for g in grupos_todos if "semillero" not in (g[2] or "").lower()]
        semilleros = [g for g in grupos_todos if "semillero" in (g[2] or "").lower()]

        if grupos_invest:
            items = "".join(
                f"<p style='margin:2px 0;'>&#9679; <b>{_s(g[2])}</b>"
                f" <span style='color:#666;'>({_s(g[4])})</span></p>"
                for g in grupos_invest
            )
            html += _seccion(
                f"Grupos de investigación ({len(grupos_invest)})",
                "#2c3e50", items
            )

        if semilleros:
            items = "".join(
                f"<p style='margin:2px 0;'>&#9679; <b>{_s(g[2])}</b>"
                f" <span style='color:#666;'>({_s(g[4])})</span></p>"
                for g in semilleros
            )
            html += _seccion(
                f"Semilleros de investigación ({len(semilleros)})",
                "#5b6b73", items
            )

        # ── Coherencia interno ↔ GrupLAC ─────────────────────────────────
        nombre_norm = normalizar_nombre(info["nombre"])
        cache_gruplac = self._obtener_cache_gruplac()
        notas_gruplac = []

        grupos_internos_norm = set()
        for g in grupos_invest:
            nombre_grupo = g[2]
            grupo_resuelto = self._analisis_gruplac.resolver_grupo_en_gruplac(nombre_grupo, cache_gruplac)
            grupos_internos_norm.add(self._analisis_gruplac._normalizar_grupo(nombre_grupo))
            if grupo_resuelto is not None:
                grupos_internos_norm.add(grupo_resuelto)

            if grupo_resuelto is None:
                # No hay ningún reporte GrupLAC para este grupo: no se puede
                # confirmar ni desmentir nada, así que no se muestra advertencia.
                continue

            entry = self._analisis_gruplac._nombre_en_gruplac(nombre_norm, nombre_grupo, cache_gruplac)
            if entry is None:
                notas_gruplac.append((
                    "warn",
                    f"No aparece en GrupLAC para el grupo <b>{_s(nombre_grupo)}</b>. Está "
                    "registrado en la base de datos interna de la Vicerrectoría de "
                    "Investigación, Innovación y Extensión, pero no en GrupLAC — favor "
                    "revisar o actualizar."
                ))
            else:
                if entry["activo"]:
                    estado = f"activo desde {entry['desde']}"
                elif entry["hasta"] != "—":
                    estado = f"retirado el {entry['hasta']} (desde {entry['desde']})"
                else:
                    estado = "vinculación registrada en GrupLAC"
                notas_gruplac.append((
                    "ok",
                    f"Sí aparece en GrupLAC para <b>{_s(nombre_grupo)}</b> ({estado})."
                ))

        hallazgos_extra = self._analisis_gruplac.buscar_persona_en_gruplac(
            nombre_norm, cache_gruplac, excluir_grupos_norm=grupos_internos_norm
        )
        for grupo_norm, entry in hallazgos_extra:
            notas_gruplac.append((
                "warn",
                f"Aparece en GrupLAC en el grupo <b>{_s(grupo_norm.title())}</b>, pero no "
                "está registrado en ese grupo en la base de datos interna — favor revisar "
                "o actualizar para que ambos sistemas coincidan."
            ))

        if notas_gruplac:
            items = "".join(
                f"<p style='margin:3px 0;padding:4px 6px;border-radius:4px;"
                f"background:{'#eafaf1' if tipo == 'ok' else '#fdf2e3'};"
                f"color:{'#1e7e34' if tipo == 'ok' else '#9a6700'};'>"
                f"{'✓' if tipo == 'ok' else '⚠'} {texto}</p>"
                for tipo, texto in notas_gruplac
            )
            html += _seccion("Coherencia con GrupLAC", "#6c5ce7", items)

        if detalle.get("publicaciones"):
            items = "".join(
                f"<p style='margin:2px 0;border-bottom:1px solid #eee;padding-bottom:2px;'>"
                f"&#9679; <b>{_s(pub[2])[:120]}</b><br>"
                f"<span style='color:#555;'>Año: {_s(pub[5])} | Tipo: {_s(pub[6])} | "
                f"Revista: {_s(pub[3])[:60]}</span>{_extra_html(pub)}</p>"
                for pub in detalle["publicaciones"]
            )
            html += _seccion(
                f"Publicaciones ({len(detalle['publicaciones'])})",
                "#2e86ab", items
            )

        if detalle.get("extensiones"):
            items = "".join(
                f"<p style='margin:2px 0;'>&#9679; <b>{_s(ext[2])[:100]}</b> "
                f"<span style='color:#555;'>({_s(ext[3])} · {_s(ext[8])})</span>"
                f"{_extra_html(ext)}</p>"
                for ext in detalle["extensiones"]
            )
            html += _seccion(
                f"Extensiones ({len(detalle['extensiones'])})",
                "#f18f01", items
            )

        if detalle.get("trabajos_grado_director"):
            items = "".join(
                f"<p style='margin:2px 0;'>&#9679; <b>{_s(tg[5])[:100]}</b><br>"
                f"<span style='color:#555;'>Estudiante: {_s(tg[4])} | Prog: {_s(tg[6])} | "
                f"Año: {_s(tg[7])}</span></p>"
                for tg in detalle["trabajos_grado_director"]
            )
            html += _seccion(
                f"Trabajos de grado dirigidos ({len(detalle['trabajos_grado_director'])})",
                "#3b8c66", items
            )

        if detalle.get("trabajos_grado_estudiante"):
            items = "".join(
                f"<p style='margin:2px 0;'>&#9679; <b>{_s(tg[5])[:100]}</b><br>"
                f"<span style='color:#555;'>Director: {_s(tg[2])} | Prog: {_s(tg[6])} | "
                f"Año: {_s(tg[7])}</span></p>"
                for tg in detalle["trabajos_grado_estudiante"]
            )
            html += _seccion(
                f"Trabajos de grado (como estudiante) ({len(detalle['trabajos_grado_estudiante'])})",
                "#52796f", items
            )

        if detalle.get("productos_innovacion"):
            items = "".join(
                f"<p style='margin:2px 0;'>&#9679; <b>{_s(p[3])[:100]}</b> "
                f"<span style='color:#555;'>({_s(p[2])} · {_s(p[5])})</span>"
                f"{_extra_html(p)}</p>"
                for p in detalle["productos_innovacion"]
            )
            html += _seccion(
                f"Productos de innovación ({len(detalle['productos_innovacion'])})",
                "#a23b72", items
            )

        if detalle.get("proyectos"):
            items = "".join(
                f"<p style='margin:2px 0;'>&#9679; <b>{_s(pr[3])[:100]}</b> "
                f"<span style='color:#555;'>(Año: {_s(pr[7])} · {_s(pr[10])})</span></p>"
                for pr in detalle["proyectos"]
            )
            html += _seccion(
                f"Proyectos ({len(detalle['proyectos'])})",
                "#c73e1d", items
            )

        html += "</body></html>"
        self.texto_detalle.setHtml(html)


class VistaGrupos(QWidget):
    def __init__(self, db):
        super().__init__()
        self.db = db
        self.productos_completos = []
        self._analisis_gruplac = AnalisisDuplicados(db.conn)
        self._cache_gruplac = None
        self._cache_notas_autor = {}
        self.setup_ui()

    def _obtener_cache_gruplac(self):
        if self._cache_gruplac is None:
            self._cache_gruplac = self._analisis_gruplac._cargar_integrantes_gruplac()
        return self._cache_gruplac
    
    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(2)
        layout.setContentsMargins(3, 3, 3, 3)
        
        # ===== BARRA SUPERIOR SIN ESTADÍSTICAS =====
        barra_superior = QWidget()
        barra_superior.setMaximumHeight(35)
        barra_superior.setStyleSheet("background-color: #f8f9fa; border-radius: 3px;")
        layout_barra = QHBoxLayout(barra_superior)
        layout_barra.setContentsMargins(5, 3, 5, 3)
        layout_barra.setSpacing(5)
        
        # Grupo
        layout_barra.addWidget(QLabel("<b>Grupo:</b>"))
        self.combo_grupos = QComboBox()
        self.combo_grupos.setMinimumWidth(200)
        self.combo_grupos.currentTextChanged.connect(self.seleccionar_grupo)
        layout_barra.addWidget(self.combo_grupos)
        
        # Botón refrescar
        btn_refrescar = QPushButton("↻")
        btn_refrescar.setMaximumWidth(30)
        btn_refrescar.setToolTip("Refrescar grupos")
        btn_refrescar.clicked.connect(self.cargar_grupos)
        layout_barra.addWidget(btn_refrescar)
        
        layout_barra.addWidget(QLabel("|"))
        
        # Filtros inline
        self.check_pub = QCheckBox("Pub")
        self.check_pub.setChecked(True)
        self.check_pub.stateChanged.connect(self.seleccionar_grupo)
        layout_barra.addWidget(self.check_pub)
        
        self.check_ext = QCheckBox("Ext")
        self.check_ext.setChecked(True)
        self.check_ext.stateChanged.connect(self.seleccionar_grupo)
        layout_barra.addWidget(self.check_ext)
        
        self.check_tg = QCheckBox("TG")
        self.check_tg.setChecked(True)
        self.check_tg.stateChanged.connect(self.seleccionar_grupo)
        layout_barra.addWidget(self.check_tg)
        
        self.check_innov = QCheckBox("Inn")
        self.check_innov.setChecked(True)
        self.check_innov.stateChanged.connect(self.seleccionar_grupo)
        layout_barra.addWidget(self.check_innov)
        
        self.check_proy = QCheckBox("Proy")
        self.check_proy.setChecked(True)
        self.check_proy.stateChanged.connect(self.seleccionar_grupo)
        layout_barra.addWidget(self.check_proy)
        
        layout_barra.addStretch()
        
        # Botones exportar
        btn_excel = QPushButton("Excel")
        btn_excel.setMaximumWidth(70)
        btn_excel.clicked.connect(self.exportar_excel)
        btn_excel.setStyleSheet("background-color: #27ae60; color: white; font-size: 11px;")
        layout_barra.addWidget(btn_excel)
        
        btn_pdf = QPushButton("PDF")
        btn_pdf.setMaximumWidth(60)
        btn_pdf.clicked.connect(self.exportar_pdf)
        btn_pdf.setStyleSheet("background-color: #e74c3c; color: white; font-size: 11px;")
        layout_barra.addWidget(btn_pdf)
        
        # ===== PANEL PRINCIPAL - 3 COLUMNAS =====
        splitter = QSplitter(Qt.Horizontal)
        
        # COLUMNA 1: Integrantes (20%)
        panel_integrantes = QWidget()
        layout_int = QVBoxLayout(panel_integrantes)
        layout_int.setSpacing(2)
        layout_int.setContentsMargins(2, 2, 2, 2)
        
        lbl_int = QLabel("<b>Integrantes</b>")
        lbl_int.setStyleSheet("color: #1a365d; font-size: 12px;")
        layout_int.addWidget(lbl_int)
        
        self.tabla_integrantes = QTableWidget()
        self.tabla_integrantes.setColumnCount(4)
        self.tabla_integrantes.setHorizontalHeaderLabels(['Nombre', 'Tipo', 'Email', 'Facultad'])
        self.tabla_integrantes.horizontalHeader().setStretchLastSection(True)
        self.tabla_integrantes.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tabla_integrantes.verticalHeader().setVisible(False)
        self.tabla_integrantes.setAlternatingRowColors(True)
        self.tabla_integrantes.setStyleSheet("font-size: 10px;")
        layout_int.addWidget(self.tabla_integrantes)
        
        # COLUMNA 2: Productos (55%)
        panel_productos = QWidget()
        layout_prod = QVBoxLayout(panel_productos)
        layout_prod.setSpacing(2)
        layout_prod.setContentsMargins(2, 2, 2, 2)
        
        lbl_prod = QLabel("<b>Productos del Grupo</b>")
        lbl_prod.setStyleSheet("color: #1a365d; font-size: 12px;")
        layout_prod.addWidget(lbl_prod)
        
        self.tabla_productos = QTableWidget()
        self.tabla_productos.setColumnCount(6)
        self.tabla_productos.setHorizontalHeaderLabels([
            'Investigador', 'Título', 'Año', 'Tipo', 'Categoría', 'Estado'
        ])
        self.tabla_productos.horizontalHeader().setStretchLastSection(True)
        self.tabla_productos.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tabla_productos.verticalHeader().setVisible(False)
        self.tabla_productos.setAlternatingRowColors(True)
        self.tabla_productos.setStyleSheet("font-size: 10px;")
        self.tabla_productos.clicked.connect(self.mostrar_detalle_producto)
        layout_prod.addWidget(self.tabla_productos)
        
        # COLUMNA 3: Detalle del Producto (25%)
        panel_detalle = QWidget()
        layout_det = QVBoxLayout(panel_detalle)
        layout_det.setSpacing(2)
        layout_det.setContentsMargins(2, 2, 2, 2)
        
        lbl_det = QLabel("<b>Detalle del Producto</b>")
        lbl_det.setStyleSheet("color: #1a365d; font-size: 12px;")
        layout_det.addWidget(lbl_det)
        
        self.texto_detalle = QTextEdit()
        self.texto_detalle.setReadOnly(True)
        self.texto_detalle.setPlaceholderText("Selecciona un producto para ver detalles...")
        self.texto_detalle.setStyleSheet("font-size: 10px; background-color: #fafafa;")
        layout_det.addWidget(self.texto_detalle)
        
        # Agregar columnas al splitter
        splitter.addWidget(panel_integrantes)
        splitter.addWidget(panel_productos)
        splitter.addWidget(panel_detalle)
        splitter.setSizes([280, 750, 370])
        
        # Agregar todo
        layout.addWidget(barra_superior)
        layout.addWidget(splitter)
        
        self.setLayout(layout)
        self.cargar_grupos()
    
    def cargar_grupos(self):
        self.combo_grupos.clear()
        self.combo_grupos.addItem("-- Seleccionar --")
        grupos = self.db.obtener_grupos()
        for grupo in grupos:
            if grupo[0]:
                self.combo_grupos.addItem(grupo[0])
    
    def seleccionar_grupo(self, grupo=None):
        if grupo is None:
            grupo = self.combo_grupos.currentText()
            
        if grupo == "-- Seleccionar --" or not grupo:
            self.tabla_integrantes.setRowCount(0)
            self.tabla_productos.setRowCount(0)
            self.texto_detalle.clear()
            return
        
        # Cargar integrantes
        integrantes = self.db.obtener_integrantes_grupo(grupo)
        self.tabla_integrantes.setRowCount(len(integrantes))
        
        for row, persona in enumerate(integrantes):
            self.tabla_integrantes.setItem(row, 0, QTableWidgetItem(persona[1]))
            self.tabla_integrantes.setItem(row, 1, QTableWidgetItem(persona[2] or ''))
            self.tabla_integrantes.setItem(row, 2, QTableWidgetItem(persona[3] or ''))
            self.tabla_integrantes.setItem(row, 3, QTableWidgetItem(persona[4] or ''))
        
        # Generar reporte automáticamente
        filtros = []
        if self.check_pub.isChecked():
            filtros.append('Publicaciones')
        if self.check_ext.isChecked():
            filtros.append('Extensiones')
        if self.check_tg.isChecked():
            filtros.append('Trabajos de Grado')
        if self.check_innov.isChecked():
            filtros.append('Productos Innovación')
        if self.check_proy.isChecked():
            filtros.append('Proyectos')
        
        if not filtros:
            self.tabla_productos.setRowCount(0)
            return
        
        productos = self.db.obtener_productos_grupo_detallado(grupo, filtros)
        productos = self._agrupar_por_titulo(productos)
        self.productos_completos = productos
        self.tabla_productos.setRowCount(len(productos))
        
        for row, prod in enumerate(productos):
            item_inv = QTableWidgetItem(prod.get('investigador', ''))
            item_inv.setData(Qt.UserRole, prod)
            self.tabla_productos.setItem(row, 0, item_inv)
            
            self.tabla_productos.setItem(row, 1, QTableWidgetItem(prod.get('titulo', '')))
            self.tabla_productos.setItem(row, 2, QTableWidgetItem(str(prod.get('año', '')) if prod.get('año') else ''))
            self.tabla_productos.setItem(row, 3, QTableWidgetItem(prod.get('tipo_producto', '')))
            self.tabla_productos.setItem(row, 4, QTableWidgetItem(prod.get('categoria', '')))
            self.tabla_productos.setItem(row, 5, QTableWidgetItem(prod.get('estado', '')))

    def _agrupar_por_titulo(self, productos):
        """
        Un mismo producto con varios coautores del mismo grupo aparece una vez
        por autor en `productos` (una fila por cédula en la tabla origen).
        Esto los agrupa en una sola entrada por (tipo, título, año), listando
        todos los autores juntos en 'investigador' en vez de repetir el
        producto — así se distingue un producto realmente repetido de uno
        que solo tiene varios autores.
        """
        grupos = {}
        orden = []
        for prod in productos:
            clave = (prod.get('tipo_producto'), normalizar_nombre(prod.get('titulo') or ''), prod.get('año'))
            if clave not in grupos:
                nuevo = dict(prod)
                nuevo['autores'] = []
                grupos[clave] = nuevo
                orden.append(clave)
            nombre = prod.get('investigador')
            autores = grupos[clave]['autores']
            if nombre and not any(a['nombre'] == nombre for a in autores):
                autores.append({
                    'nombre': nombre,
                    'cedula': prod.get('cedula'),
                    'nota_gruplac': self._nota_gruplac_autor(nombre, prod.get('cedula')),
                })

        resultado = []
        for clave in orden:
            g = grupos[clave]
            g['investigador'] = '; '.join(self._formatear_autor(a) for a in g['autores'])
            resultado.append(g)
        return resultado

    @staticmethod
    def _formatear_autor(autor):
        nombre = autor.get('nombre') or ''
        nota = autor.get('nota_gruplac')
        return f"{nombre} {nota}" if nota else nombre

    def _nota_gruplac_autor(self, nombre, cedula):
        """
        Si esta persona también aparece en GrupLAC bajo un grupo donde NO
        está registrada internamente, devuelve una nota corta para mostrar
        junto a su nombre. Reutiliza la misma lógica de "Búsqueda de
        Personas": excluye TODOS los grupos internos de la persona (no solo
        el grupo que se está reportando) — si solo excluyéramos el grupo
        actual, alguien correctamente registrado en GrupLAC bajo OTRO de sus
        propios grupos internos quedaría marcado como inconsistencia falsa
        (esto pasaba con Mauricio Holguín: está en 4 grupos internos, y el
        reporte de uno de ellos lo marcaba mal por no conocer los otros 3).
        """
        if not nombre:
            return ''
        nombre_norm = normalizar_nombre(nombre)
        clave_cache = (nombre_norm, cedula)
        if clave_cache in self._cache_notas_autor:
            return self._cache_notas_autor[clave_cache]

        cache = self._obtener_cache_gruplac()
        grupos_persona = []
        if cedula:
            grupos_persona = [
                r[0] for r in self.db.conn.execute(
                    "SELECT DISTINCT grupo FROM grupos WHERE cedula = ? "
                    "AND grupo IS NOT NULL AND grupo != ''",
                    (cedula,)
                ).fetchall()
            ]

        grupos_internos_norm = set()
        for g in grupos_persona:
            grupos_internos_norm.add(self._analisis_gruplac._normalizar_grupo(g))
            resuelto = self._analisis_gruplac.resolver_grupo_en_gruplac(g, cache)
            if resuelto is not None:
                grupos_internos_norm.add(resuelto)

        hallazgos = self._analisis_gruplac.buscar_persona_en_gruplac(
            nombre_norm, cache, excluir_grupos_norm=grupos_internos_norm
        )
        if not hallazgos:
            nota = ''
        else:
            grupos_extra = ', '.join(sorted({g.title() for g, _ in hallazgos}))
            nota = (
                f"⚠ también aparece en GrupLAC en el grupo {grupos_extra}, "
                "donde no está registrado en la BD interna"
            )

        self._cache_notas_autor[clave_cache] = nota
        return nota

    def mostrar_detalle_producto(self, index):
        row = index.row()
        item = self.tabla_productos.item(row, 0)
        if not item:
            return

        producto = item.data(Qt.UserRole)
        if not producto:
            return

        tipo = producto.get('tipo_producto', '')

        COLORES_TIPO = {
            'Publicación': '#2e86ab',
            'Extensión': '#f18f01',
            'Trabajo de Grado': '#3b8c66',
            'Innovación': '#a23b72',
            'Proyecto': '#c73e1d',
        }
        color = COLORES_TIPO.get(tipo, '#1a365d')

        def _s(v):
            return str(v).strip() if v and str(v).strip() not in ('', 'None', 'nan') else '—'

        def _fila(label, valor):
            return (
                f"<tr><td style='color:#555;white-space:nowrap;padding:2px 6px 2px 0;"
                f"vertical-align:top;'><b>{label}</b></td>"
                f"<td style='padding:2px 0;word-break:break-word;'>{_s(valor)}</td></tr>"
            )

        autores = producto.get('autores') or []
        label_investigador = "Investigadores" if len(autores) > 1 else "Investigador"
        fila_cedula = "" if len(autores) > 1 else _fila("Cédula", producto.get('cedula'))

        html = (
            "<html><body style='font-family:Arial,sans-serif;font-size:11px;"
            "max-width:100%;word-wrap:break-word;margin:0;padding:4px;'>"
            f"<div style='background:{color};color:white;padding:6px 10px;"
            f"border-radius:5px;margin-bottom:8px;'>"
            f"<b style='font-size:12px;'>{tipo.upper()}</b></div>"
            "<table style='width:100%;border-collapse:collapse;'>"
            + _fila(label_investigador, producto.get('investigador'))
            + fila_cedula
            + _fila("Año", producto.get('año'))
            + "</table>"
            f"<div style='background:{color};color:white;padding:3px 8px;"
            f"border-radius:3px;margin:8px 0 4px 0;font-size:10px;"
            f"font-weight:bold;'>INFORMACIÓN DETALLADA</div>"
            "<table style='width:100%;border-collapse:collapse;'>"
        )

        if tipo == 'Publicación':
            html += (
                f"<tr><td colspan='2' style='padding:2px 0;word-break:break-word;'>"
                f"<b>Título:</b> {_s(producto.get('titulo'))}</td></tr>"
                f"<tr><td colspan='2' style='padding:2px 0;word-break:break-word;"
                f"color:#444;'><b>Revista/Libro:</b> {_s(producto.get('revista_libro'))}</td></tr>"
                + _fila("DOI/URL", producto.get('doi_url'))
                + _fila("ISSN/ISBN", producto.get('issn_isbn'))
                + _fila("Categoría", producto.get('categoria'))
                + _fila("Estado", producto.get('estado'))
            )

        elif tipo == 'Extensión':
            html += (
                f"<tr><td colspan='2' style='padding:2px 0;word-break:break-word;'>"
                f"<b>Actividad:</b> {_s(producto.get('titulo'))}</td></tr>"
                + _fila("Tipo", producto.get('tipo'))
                + _fila("Modalidad", producto.get('modalidad'))
                + _fila("Estado", producto.get('estado'))
                + _fila("Fechas", f"{_s(producto.get('fecha_inicio'))} – {_s(producto.get('fecha_fin'))}")
                + _fila("Población", producto.get('poblacion'))
                + _fila("Financiación interna", producto.get('financiacion_interna'))
                + _fila("Financiación externa", producto.get('financiacion_externa'))
            )

        elif tipo == 'Trabajo de Grado':
            html += (
                f"<tr><td colspan='2' style='padding:2px 0;word-break:break-word;'>"
                f"<b>Título:</b> {_s(producto.get('titulo'))}</td></tr>"
                + _fila("Director", producto.get('investigador'))
                + _fila("Estudiante", producto.get('estudiante'))
                + _fila("Cédula estudiante", producto.get('cedula_estudiante'))
                + _fila("Programa", producto.get('programa'))
                + _fila("Estado", producto.get('estado'))
                + _fila("Fecha sustentación", producto.get('fecha_sustentacion'))
                + _fila("Calificación", producto.get('calificacion'))
            )

        elif tipo == 'Innovación':
            html += (
                f"<tr><td colspan='2' style='padding:2px 0;word-break:break-word;'>"
                f"<b>Nombre:</b> {_s(producto.get('titulo'))}</td></tr>"
                + _fila("Tipo detalle", producto.get('tipo_producto_detalle'))
                + f"<tr><td colspan='2' style='padding:2px 0;word-break:break-word;'>"
                f"<b>Descripción:</b> {_s(producto.get('descripcion'))}</td></tr>"
                + _fila("Estado", producto.get('estado'))
                + _fila("Grupo", producto.get('grupo'))
            )

        elif tipo == 'Proyecto':
            html += (
                f"<tr><td colspan='2' style='padding:2px 0;word-break:break-word;'>"
                f"<b>Título:</b> {_s(producto.get('titulo'))}</td></tr>"
                f"<tr><td colspan='2' style='padding:2px 0;word-break:break-word;"
                f"color:#444;'><b>Objetivo:</b> {_s(producto.get('objetivo'))}</td></tr>"
                + _fila("Código CIE", producto.get('codigo_cie'))
                + _fila("Tipo", producto.get('tipo'))
                + _fila("Fechas", f"{_s(producto.get('fecha_inicio'))} – {_s(producto.get('fecha_fin'))}")
                + _fila("Estado", producto.get('estado'))
                + _fila("Valor aprobado", producto.get('valor_aprobado'))
            )

        html += (
            "</table>"
            "<div style='margin-top:8px;padding-top:4px;border-top:1px solid #ddd;"
            f"color:#666;font-size:10px;'>Fuente: {_s(producto.get('fuente'))}</div>"
            "</body></html>"
        )

        self.texto_detalle.setHtml(html)
    
    def exportar_excel(self):
        grupo = self.combo_grupos.currentText()
        if grupo == "-- Seleccionar --" or not grupo:
            QMessageBox.warning(self, "Advertencia", "Selecciona un grupo primero")
            return
        
        if not self.productos_completos:
            QMessageBox.warning(self, "Advertencia", "No hay productos para exportar")
            return
        
        try:
            from datetime import datetime
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            nombre_archivo = f"reporte_{grupo.replace(' ', '_')}_{timestamp}.xlsx"
            
            # Crear workbook
            wb = openpyxl.Workbook()
            
            # ===== HOJA 1: INTEGRANTES =====
            ws_int = wb.active
            ws_int.title = "Integrantes"
            
            # Título
            ws_int['A1'] = f"GRUPO: {grupo} - INTEGRANTES"
            ws_int['A1'].font = Font(bold=True, size=14, color="1a365d")
            ws_int.merge_cells('A1:E1')
            
            # Encabezados
            encabezados_int = ['Nombre', 'Tipo', 'Cédula', 'Email', 'Facultad']
            for col, enc in enumerate(encabezados_int, 1):
                cell = ws_int.cell(row=2, column=col, value=enc)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="1a365d", end_color="1a365d", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # Datos integrantes
            integrantes = self.db.obtener_integrantes_grupo(grupo)
            for row, persona in enumerate(integrantes, 3):
                ws_int.cell(row=row, column=1, value=persona[1])
                ws_int.cell(row=row, column=2, value=persona[2] or '')
                ws_int.cell(row=row, column=3, value=persona[0])
                ws_int.cell(row=row, column=4, value=persona[3] or '')
                ws_int.cell(row=row, column=5, value=persona[4] or '')
            
            # Ajustar anchos
            for col in range(1, 6):
                ws_int.column_dimensions[get_column_letter(col)].width = 25
            
            # ===== HOJA 2: PUBLICACIONES =====
            publicaciones = [p for p in self.productos_completos if p.get('tipo_producto') == 'Publicación']
            if publicaciones:
                ws_pub = wb.create_sheet("Publicaciones")
                
                ws_pub['A1'] = f"GRUPO: {grupo} - PUBLICACIONES"
                ws_pub['A1'].font = Font(bold=True, size=14, color="1a365d")
                ws_pub.merge_cells('A1:H1')
                
                headers = ['Investigador', 'Título', 'Revista/Libro', 'Año', 'Tipo', 'Categoría', 'ISSN/ISBN', 'Estado']
                for col, h in enumerate(headers, 1):
                    cell = ws_pub.cell(row=2, column=col, value=h)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="3498db", end_color="3498db", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center")
                
                for row, pub in enumerate(publicaciones, 3):
                    ws_pub.cell(row=row, column=1, value=pub.get('investigador', ''))
                    ws_pub.cell(row=row, column=2, value=pub.get('titulo', ''))
                    ws_pub.cell(row=row, column=3, value=pub.get('revista_libro', ''))
                    ws_pub.cell(row=row, column=4, value=pub.get('año', ''))
                    ws_pub.cell(row=row, column=5, value=pub.get('tipo', ''))
                    ws_pub.cell(row=row, column=6, value=pub.get('categoria', ''))
                    ws_pub.cell(row=row, column=7, value=pub.get('issn_isbn', ''))
                    ws_pub.cell(row=row, column=8, value=pub.get('estado', ''))
                
                for col in range(1, 9):
                    ws_pub.column_dimensions[get_column_letter(col)].width = 20
                ws_pub.column_dimensions['B'].width = 50
            
            # ===== HOJA 3: EXTENSIONES =====
            extensiones = [p for p in self.productos_completos if p.get('tipo_producto') == 'Extensión']
            if extensiones:
                ws_ext = wb.create_sheet("Extensiones")
                
                ws_ext['A1'] = f"GRUPO: {grupo} - EXTENSIONES"
                ws_ext['A1'].font = Font(bold=True, size=14, color="1a365d")
                ws_ext.merge_cells('A1:G1')
                
                headers = ['Investigador', 'Actividad', 'Tipo', 'Modalidad', 'Año', 'Población', 'Estado']
                for col, h in enumerate(headers, 1):
                    cell = ws_ext.cell(row=2, column=col, value=h)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="e74c3c", end_color="e74c3c", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center")
                
                for row, ext in enumerate(extensiones, 3):
                    ws_ext.cell(row=row, column=1, value=ext.get('investigador', ''))
                    ws_ext.cell(row=row, column=2, value=ext.get('titulo', ''))
                    ws_ext.cell(row=row, column=3, value=ext.get('tipo', ''))
                    ws_ext.cell(row=row, column=4, value=ext.get('modalidad', ''))
                    ws_ext.cell(row=row, column=5, value=ext.get('año', ''))
                    ws_ext.cell(row=row, column=6, value=ext.get('poblacion', ''))
                    ws_ext.cell(row=row, column=7, value=ext.get('estado', ''))
                
                for col in range(1, 8):
                    ws_ext.column_dimensions[get_column_letter(col)].width = 20
                ws_ext.column_dimensions['B'].width = 50
            
            # ===== HOJA 4: TRABAJOS DE GRADO =====
            trabajos = [p for p in self.productos_completos if p.get('tipo_producto') == 'Trabajo de Grado']
            if trabajos:
                ws_tg = wb.create_sheet("Trabajos de Grado")
                
                ws_tg['A1'] = f"GRUPO: {grupo} - TRABAJOS DE GRADO"
                ws_tg['A1'].font = Font(bold=True, size=14, color="1a365d")
                ws_tg.merge_cells('A1:G1')
                
                headers = ['Director', 'Título', 'Estudiante', 'Programa', 'Año', 'Estado', 'Calificación']
                for col, h in enumerate(headers, 1):
                    cell = ws_tg.cell(row=2, column=col, value=h)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="9b59b6", end_color="9b59b6", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center")
                
                for row, tg in enumerate(trabajos, 3):
                    ws_tg.cell(row=row, column=1, value=tg.get('investigador', ''))
                    ws_tg.cell(row=row, column=2, value=tg.get('titulo', ''))
                    ws_tg.cell(row=row, column=3, value=tg.get('estudiante', ''))
                    ws_tg.cell(row=row, column=4, value=tg.get('programa', ''))
                    ws_tg.cell(row=row, column=5, value=tg.get('año', ''))
                    ws_tg.cell(row=row, column=6, value=tg.get('estado', ''))
                    ws_tg.cell(row=row, column=7, value=tg.get('calificacion', ''))
                
                for col in range(1, 8):
                    ws_tg.column_dimensions[get_column_letter(col)].width = 20
                ws_tg.column_dimensions['B'].width = 50
            
            # ===== HOJA 5: PROYECTOS =====
            proyectos = [p for p in self.productos_completos if p.get('tipo_producto') == 'Proyecto']
            if proyectos:
                ws_proy = wb.create_sheet("Proyectos")
                
                ws_proy['A1'] = f"GRUPO: {grupo} - PROYECTOS"
                ws_proy['A1'].font = Font(bold=True, size=14, color="1a365d")
                ws_proy.merge_cells('A1:G1')
                
                headers = ['Investigador', 'Título', 'Código CIE', 'Tipo', 'Año', 'Estado', 'Valor']
                for col, h in enumerate(headers, 1):
                    cell = ws_proy.cell(row=2, column=col, value=h)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="1abc9c", end_color="1abc9c", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center")
                
                for row, proy in enumerate(proyectos, 3):
                    ws_proy.cell(row=row, column=1, value=proy.get('investigador', ''))
                    ws_proy.cell(row=row, column=2, value=proy.get('titulo', ''))
                    ws_proy.cell(row=row, column=3, value=proy.get('codigo_cie', ''))
                    ws_proy.cell(row=row, column=4, value=proy.get('tipo', ''))
                    ws_proy.cell(row=row, column=5, value=proy.get('año', ''))
                    ws_proy.cell(row=row, column=6, value=proy.get('estado', ''))
                    ws_proy.cell(row=row, column=7, value=proy.get('valor_aprobado', ''))
                
                for col in range(1, 8):
                    ws_proy.column_dimensions[get_column_letter(col)].width = 20
                ws_proy.column_dimensions['B'].width = 50
            
            # ===== HOJA 6: INNOVACIÓN =====
            innovacion = [p for p in self.productos_completos if p.get('tipo_producto') == 'Innovación']
            if innovacion:
                ws_inn = wb.create_sheet("Innovación")
                
                ws_inn['A1'] = f"GRUPO: {grupo} - PRODUCTOS DE INNOVACIÓN"
                ws_inn['A1'].font = Font(bold=True, size=14, color="1a365d")
                ws_inn.merge_cells('A1:F1')
                
                headers = ['Investigador', 'Nombre', 'Tipo', 'Año', 'Estado', 'Descripción']
                for col, h in enumerate(headers, 1):
                    cell = ws_inn.cell(row=2, column=col, value=h)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="f39c12", end_color="f39c12", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center")
                
                for row, inn in enumerate(innovacion, 3):
                    ws_inn.cell(row=row, column=1, value=inn.get('investigador', ''))
                    ws_inn.cell(row=row, column=2, value=inn.get('titulo', ''))
                    ws_inn.cell(row=row, column=3, value=inn.get('tipo_producto_detalle', ''))
                    ws_inn.cell(row=row, column=4, value=inn.get('año', ''))
                    ws_inn.cell(row=row, column=5, value=inn.get('estado', ''))
                    ws_inn.cell(row=row, column=6, value=inn.get('descripcion', ''))
                
                for col in range(1, 7):
                    ws_inn.column_dimensions[get_column_letter(col)].width = 20
                ws_inn.column_dimensions['B'].width = 40
                ws_inn.column_dimensions['F'].width = 60
            
            # Guardar
            wb.save(nombre_archivo)
            
            QMessageBox.information(self, "Éxito", 
                f"Excel exportado exitosamente:\n{nombre_archivo}\n\n"
                f"Integrantes: {len(integrantes)}\n"
                f"Productos totales: {len(self.productos_completos)}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al exportar Excel:\n{str(e)}")
    
    def exportar_pdf(self):
        """Exporta un reporte PDF del grupo seleccionado"""
        
        # ============= PASO 1: OBTENER NOMBRE DEL GRUPO =============
        grupo = self.combo_grupos.currentText()
        
        # ============= PASO 2: VALIDACIONES =============
        if grupo == "-- Seleccionar --" or not grupo:
            QMessageBox.warning(self, "Advertencia", "Selecciona un grupo primero")
            return
        
        if not self.productos_completos:
            QMessageBox.warning(self, "Advertencia", "No hay productos para exportar")
            return
        
        try:
            # ============= PASO 3: IMPORTACIONES =============
            from datetime import datetime
            from PyQt5.QtPrintSupport import QPrinter
            from PyQt5.QtGui import QTextDocument
            
            # ============= PASO 4: CREAR TIMESTAMP =============
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # ============= PASO 5: LIMPIAR NOMBRE DEL GRUPO =============
            # Opción 1: Limpieza básica (recomendada, más simple)
            grupo_limpio = grupo.replace(' ', '_').replace(':', '-').replace('/', '-')
            
            # Opción 2: Limpieza completa (si agregaste la función)
            # grupo_limpio = limpiar_nombre_archivo(grupo)
            
            # ============= PASO 6: CREAR NOMBRE DE ARCHIVO =============
            nombre_archivo = f"reporte_{grupo_limpio}_{timestamp}.pdf"
            
            # ============= PASO 7: CREAR DOCUMENTO HTML =============
            html = f"""
            <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 20px; }}
                    h1 {{ color: #1a365d; border-bottom: 3px solid #1a365d; padding-bottom: 10px; }}
                    h2 {{ color: #2c3e50; margin-top: 30px; border-bottom: 2px solid #ecf0f1; padding-bottom: 5px; }}
                    table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 9pt; }}
                    th {{ background-color: #1a365d; color: white; padding: 8px; text-align: left; }}
                    td {{ border: 1px solid #ddd; padding: 6px; }}
                    tr:nth-child(even) {{ background-color: #f2f2f2; }}
                    .producto {{ margin-bottom: 15px; padding: 10px; border: 1px solid #ddd; border-radius: 5px; page-break-inside: avoid; }}
                </style>
            </head>
            <body>
                <h1>Reporte del Grupo: {grupo}</h1>
                <p><strong>Fecha:</strong> {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}</p>
            """
            
            # ============= PASO 8: INTEGRANTES =============
            integrantes = self.db.obtener_integrantes_grupo(grupo)
            html += f"""
                <h2>Integrantes ({len(integrantes)})</h2>
                <table>
                    <tr><th>Nombre</th><th>Tipo</th><th>Cédula</th><th>Email</th></tr>
            """
            
            for persona in integrantes:
                html += f"""
                    <tr>
                        <td>{persona[1]}</td>
                        <td>{persona[2] or ''}</td>
                        <td>{persona[0]}</td>
                        <td>{persona[3] or ''}</td>
                    </tr>
                """
            
            html += "</table>"
            
            # ============= PASO 9: PRODUCTOS POR TIPO =============
            tipos_plural = {
                'Publicación': 'Publicaciones',
                'Extensión': 'Extensiones',
                'Trabajo de Grado': 'Trabajos de Grado',
                'Innovación': 'Innovación',
                'Proyecto': 'Proyectos',
            }
            for tipo, plural in tipos_plural.items():
                productos_tipo = [p for p in self.productos_completos if p.get('tipo_producto') == tipo]

                if not productos_tipo:
                    continue

                html += f'<h2>{plural} ({len(productos_tipo)})</h2>'
                
                for i, prod in enumerate(productos_tipo, 1):
                    html += f'<div class="producto">'
                    html += f'<strong>#{i}</strong> - {prod.get("titulo", "Sin título")}<br>'
                    if tipo == 'Trabajo de Grado':
                        html += f'<em>Director:</em> {prod.get("investigador", "N/A")}'
                        html += f' | <em>Estudiante:</em> {prod.get("estudiante", "N/A")}'
                    else:
                        html += f'<em>Investigador:</em> {prod.get("investigador", "N/A")}'
                    if prod.get('año'):
                        html += f' | <em>Año:</em> {prod["año"]}'
                    html += '</div>'
            
            html += "</body></html>"
            
            # ============= PASO 10: CREAR PDF =============
            documento = QTextDocument()
            documento.setHtml(html)
            
            printer = QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(nombre_archivo)
            printer.setPageSize(QPrinter.A4)
            printer.setPageMargins(10, 10, 10, 10, QPrinter.Millimeter)
            
            documento.print_(printer)
            
            # ============= PASO 11: MENSAJE DE ÉXITO =============
            QMessageBox.information(self, "Éxito", f"PDF exportado:\n{nombre_archivo}")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al exportar PDF:\n{str(e)}")


class VentanaPrincipal(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = DatabaseManager()
        self._tab_registry = {}
        self._loaded_tabs = set()
        self.vista_inicio = None
        self.vista_busqueda = None
        self.vista_grupos = None
        self.setup_ui()
        self.cargar_datos_automaticamente()

    def _crear_tab(self, index):
        if index in self._loaded_tabs:
            return
        self._loaded_tabs.add(index)
        factory, args = self._tab_registry.get(index)
        if factory is None:
            return
        widget = factory(*args)
        self.tabs.blockSignals(True)
        self.tabs.removeTab(index)
        self.tabs.insertTab(index, widget, self._tab_titles[index])
        self.tabs.setCurrentIndex(index)
        self.tabs.blockSignals(False)
        if index == 0:
            self.vista_inicio = widget
            self.vista_inicio.procesar_solicitado.connect(self.cargar_datos_automaticamente)
        elif index == 2:
            self.vista_grupos = widget
        elif index == 1:
            self.vista_busqueda = widget

    def setup_ui(self):
        self.setWindowTitle("Consolidado de Información")
        self.setGeometry(100, 100, 1400, 800)
        
        widget_central = QWidget()
        self.setCentralWidget(widget_central)
        
        layout = QVBoxLayout(widget_central)
        
        header = QLabel("Consolidado de Información")
        header.setFont(QFont("Arial", 16, QFont.Bold))
        header.setStyleSheet("background-color: #1a365d; color: white; padding: 15px;")
        
        self.tabs = QTabWidget()
        self._tab_titles = [
            "Inicio",
            "Búsqueda de Personas",
            "Reportes por Grupo",
            "Seguimiento Grupos",
        ]
        self._tab_registry = {
            0: (lambda db: VistaInicio(db), (self.db,)),
            1: (lambda db: VistaBusqueda(db), (self.db,)),
            2: (lambda db: VistaGrupos(db), (self.db,)),
            3: (lambda db: VistaSeguimientoGrupos(db), (self.db,)),
        }

        for i, title in enumerate(self._tab_titles):
            self.tabs.addTab(QWidget(), title)

        self.tabs.currentChanged.connect(self._crear_tab)
        self._crear_tab(0)

        self.statusBar().showMessage("Listo")
        
        layout.addWidget(header)
        layout.addWidget(self.tabs)
    
    def cargar_datos_automaticamente(self):
        if getattr(self, "cargador", None) is not None and self.cargador.isRunning():
            return
        self.statusBar().showMessage("Cargando datos...")
        if self.vista_inicio is not None:
            self.vista_inicio.marcar_procesando("Procesando…")
            self.vista_inicio.btn_procesar.setEnabled(False)
        self.cargador = CargadorDatosIntegrado(self.db)
        self.cargador.progreso.connect(self.actualizar_status)
        self.cargador.finalizado.connect(self.carga_finalizada)
        self.cargador.duplicados_consolidados.connect(self.mostrar_duplicados_consolidados)
        self.cargador.start()

    def actualizar_status(self, mensaje):
        self.statusBar().showMessage(mensaje)
        if self.vista_inicio is not None:
            self.vista_inicio.marcar_procesando(mensaje)

    def mostrar_duplicados_consolidados(self, consolidaciones):
        if not consolidaciones:
            return

        mensaje = f"Se consolidaron {len(consolidaciones)} persona(s) con cédulas duplicadas:\n\n"

        for i, grupo in enumerate(consolidaciones, 1):
            nombre = grupo['nombre']
            cedula_principal = grupo['cedula_principal']
            todas_cedulas = grupo['cedulas']

            mensaje += f"{i}. {nombre}\n"
            mensaje += f"   Cédula principal: {cedula_principal}\n"
            mensaje += f"   Todas las cédulas: {', '.join(todas_cedulas)}\n\n"

        dialogo = QDialog(self)
        dialogo.setWindowTitle("Duplicados Consolidados")
        dialogo.resize(650, 500)
        layout = QVBoxLayout(dialogo)

        titulo = QLabel(f"Consolidación Completada — {len(consolidaciones)} persona(s)")
        titulo.setFont(QFont("Arial", 11, QFont.Bold))
        layout.addWidget(titulo)

        texto = QTextEdit()
        texto.setReadOnly(True)
        texto.setPlainText(mensaje)
        layout.addWidget(texto)

        btn_cerrar = QPushButton("Cerrar")
        btn_cerrar.clicked.connect(dialogo.accept)
        layout.addWidget(btn_cerrar)

        dialogo.exec_()
    
    def carga_finalizada(self, stats):
        partes = [
            f"Personas: {stats.get('personas', 0)}",
            f"Grupos: {stats.get('grupos', 0)}",
            f"Publicaciones: {stats.get('publicaciones', 0)}",
            f"Extensiones: {stats.get('extensiones', 0)}",
            f"Trabajos: {stats.get('trabajos', 0)}",
            f"Innovación: {stats.get('innovacion', 0)}",
            f"Proyectos: {stats.get('proyectos', 0)}",
            f"Propiedad: {stats.get('propiedad', 0)}",
        ]
        self.statusBar().showMessage("Carga completada — " + " | ".join(partes))
        if self.vista_inicio is not None:
            self.vista_inicio.procesamiento_finalizado(stats)
            self.vista_inicio.btn_procesar.setEnabled(True)
        self._crear_tab(2)
        if self.vista_grupos is not None:
            self.vista_grupos.cargar_grupos()

    def _generar_reporte_carga(self, stats):
        """
        Genera un Excel de resumen de productos por grupo tras la carga de datos,
        para que cada grupo sepa qué subir a GrupLAC.
        El archivo queda en el directorio base con el timestamp de la carga.
        """
        try:
            import openpyxl
            from datetime import datetime
            from openpyxl.styles import Alignment, Font, PatternFill

            base_dir = obtener_directorio_base()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            nombre = str(base_dir / f"reporte_pendientes_{timestamp}.xlsx")

            cur = self.db.conn.cursor()
            grupos = [
                r[0] for r in cur.execute(
                    "SELECT DISTINCT grupo FROM grupos "
                    "WHERE grupo IS NOT NULL AND grupo != '' "
                    "ORDER BY grupo"
                ).fetchall()
            ]

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Resumen por Grupo"

            fill_h = PatternFill(start_color="1a365d", end_color="1a365d", fill_type="solid")
            centrado = Alignment(horizontal="center")
            headers = ["Grupo", "Publicaciones", "Extensiones",
                       "Trabajos de Grado", "Innovación", "Proyectos"]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=h)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = fill_h
                cell.alignment = centrado

            tablas = ["publicaciones", "extensiones", "trabajos_grado",
                      "productos_innovacion", "proyectos"]
            for fila, grupo in enumerate(grupos, start=2):
                ws.cell(row=fila, column=1, value=grupo)
                for col_idx, tabla in enumerate(tablas, start=2):
                    try:
                        count = cur.execute(
                            f"SELECT COUNT(*) FROM {tabla} WHERE grupo = ?",
                            (grupo,),
                        ).fetchone()[0]
                    except Exception:
                        count = 0
                    cell = ws.cell(row=fila, column=col_idx, value=count)
                    cell.alignment = centrado

            ws.column_dimensions["A"].width = 55
            for letra in "BCDEF":
                ws.column_dimensions[letra].width = 16

            wb.save(nombre)
            self.statusBar().showMessage(f"Reporte de pendientes generado: {nombre}")
        except Exception as e:
            self.statusBar().showMessage(f"No se generó reporte de pendientes: {e}")

    def closeEvent(self, event):
        try:
            self.db.close()
        except Exception:
            pass
        event.accept()


def main():
    app = QApplication(sys.argv)
    ventana = VentanaPrincipal()
    ventana.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()