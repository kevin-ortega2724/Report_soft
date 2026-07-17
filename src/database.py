"""
Database manager for the consolidated research groups database.
Handles schema creation, data merging, person deduplication, and queries.
"""

import json
import sqlite3
from pathlib import Path

from utils import limpiar_cedula, normalizar_nombre, obtener_directorio_base


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

    # ── Fuentes adicionales (archivos que se acumulan sin reemplazar) ──

    def obtener_fuentes_adicionales(self, clave: str) -> list:
        """Archivos adicionales registrados para una categoría de
        ARCHIVOS_FUENTE_957 (más allá del archivo canónico), cada uno con
        su propia ruta única -- se procesan todos, ninguno reemplaza a otro."""
        row = self.conn.execute(
            "SELECT valor FROM configuracion WHERE clave = ?",
            (f"fuentes_adicionales_{clave}",),
        ).fetchone()
        if not row:
            return []
        try:
            return json.loads(row[0])
        except Exception:
            return []

    def registrar_fuente_adicional(self, clave: str, ruta: str, nombre_original: str):
        from datetime import datetime
        fuentes = self.obtener_fuentes_adicionales(clave)
        fuentes.append({
            "ruta": ruta,
            "nombre_original": nombre_original,
            "fecha_agregado": datetime.now().isoformat(timespec="seconds"),
        })
        self.conn.execute(
            "INSERT OR REPLACE INTO configuracion (clave, valor) VALUES (?, ?)",
            (f"fuentes_adicionales_{clave}", json.dumps(fuentes)),
        )
        self.conn.commit()

    # ── Cache de carga ──

    def guardar_sello_carga(self, archivos: list):
        sello = {}
        for ruta in archivos:
            p = Path(ruta)
            if p.exists():
                sello[str(p)] = p.stat().st_mtime
        self.conn.execute(
            "INSERT OR REPLACE INTO configuracion (clave, valor) VALUES ('sello_carga', ?)",
            (json.dumps(sello),),
        )
        self.conn.commit()

    def cache_valida(self, archivos_a_verificar: list) -> bool:
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
                p = Path(ruta)
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
            cursor.execute(
                'UPDATE personas SET nombre = ? WHERE cedula = ?',
                (nombre_principal, cedula_principal)
            )
            for cedula_dup in cedulas_duplicadas:
                cedula_dup = self.obtener_cedula_principal(cedula_dup)
                cursor.execute('''
                    INSERT OR REPLACE INTO cedulas_duplicadas (cedula_duplicada, cedula_principal)
                    VALUES (?, ?)
                ''', (cedula_dup, cedula_principal))
                self.mapa_cedulas[cedula_dup] = cedula_principal
                for tabla in ['grupos', 'publicaciones', 'extensiones', 'productos_innovacion', 'proyectos']:
                    try:
                        cursor.execute(f'UPDATE OR IGNORE {tabla} SET cedula = ? WHERE cedula = ?',
                                     (cedula_principal, cedula_dup))
                    except:
                        pass
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
                cursor.execute('DELETE FROM personas WHERE cedula = ?', (cedula_dup,))
            self.conn.commit()
        except sqlite3.OperationalError as e:
            try:
                self.conn.rollback()
            except:
                pass
            raise e
        except Exception as e:
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

    def _upsert_personas_batch(self, filas):
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
            es_duplicada = cursor.execute(
                'SELECT 1 FROM cedulas_duplicadas WHERE cedula_duplicada = ?',
                (cedula,)
            ).fetchone()
            if not es_duplicada:
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
        try:
            self.conn.commit()
            self.conn.close()
        except Exception as e:
            print("Error cerrando BD:", e)
