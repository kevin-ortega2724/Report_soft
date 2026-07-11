"""
Data loader - extracts data from Excel source files and loads them into the database.
"""

import json
import os
import re
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd
from unidecode import unidecode
from PyQt5.QtCore import QThread, pyqtSignal

from database import DatabaseManager
from utils import limpiar_cedula, limpiar_texto, norm_text, obtener_directorio_base


class CargadorDatosIntegrado(QThread):
    progreso = pyqtSignal(str)
    finalizado = pyqtSignal(dict)
    duplicados_consolidados = pyqtSignal(list)

    def __init__(self, db):
        super().__init__()
        self.db = db
        self.archivos_directorio = obtener_directorio_base()

    def _archivos_cambiados(self, rutas):
        try:
            row = self.db.conn.execute(
                "SELECT valor FROM configuracion WHERE clave='sello_carga'"
            ).fetchone()
            if not row:
                return set(rutas)
            sello = json.loads(row[0])
        except Exception:
            return set(rutas)
        cambiados = set()
        for ruta in rutas:
            p = Path(ruta)
            if not p.exists():
                continue
            mt = p.stat().st_mtime
            prev = sello.get(str(p))
            if prev is None or mt > prev:
                cambiados.add(ruta)
        return cambiados

    @staticmethod
    def _tablas_con_fuente():
        return ['publicaciones', 'extensiones', 'trabajos_grado',
                'productos_innovacion', 'proyectos', 'propiedad_intelectual']

    def _delete_datos_para_archivo(self, ruta):
        nombre = os.path.basename(ruta)
        conn = self.db.conn
        for tabla in self._tablas_con_fuente():
            conn.execute(f"DELETE FROM {tabla} WHERE fuente LIKE ?", (f'%{nombre}%',))

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
        data_input = base / "data" / "input"
        archivos_fuente = [str(data_input / n) for n in nomina_nombres]
        # Also try base dir for backwards compat
        for i, n in enumerate(nomina_nombres):
            p = Path(archivos_fuente[i])
            if not p.exists():
                alt = base / n
                if alt.exists():
                    archivos_fuente[i] = str(alt)

        if self.db.cache_valida(archivos_fuente):
            self.progreso.emit("Base de datos en caché — omitiendo reproceso")
            self.finalizado.emit(self.db.obtener_estadisticas())
            return

        conn = self.db.conn
        conn.execute("PRAGMA synchronous=OFF")
        try:
            cambiados = self._archivos_cambiados(archivos_fuente)
            primer_cambio = not cambiados or cambiados == set(archivos_fuente)

            if primer_cambio:
                self.db.limpiar_datos()
            else:
                for archivo in cambiados:
                    self._delete_datos_para_archivo(archivo)
            conn.commit()

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
                        '''INSERT INTO extensiones
                           (cedula,actividad,tipo,modalidad,estado,fecha_inicio,fecha_fin,
                            año,poblacion,grupo,facultad,financiacion_interna,financiacion_externa,fuente)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                        ext_b,
                    )
                    self.progreso.emit(f"Insertadas {len(ext_b)} extensiones")

            if resultados.get('produccion'):
                _, pub_b = resultados['produccion']
                if pub_b:
                    conn.executemany(
                        '''INSERT INTO publicaciones
                           (cedula,titulo,revista_libro,doi_url,issn_isbn,año,
                            tipo,categoria,estado,grupo,fuente)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?)''',
                        pub_b,
                    )
                    self.progreso.emit(f"Insertadas {len(pub_b)} publicaciones")

            if resultados.get('trabajos_grado'):
                _, tg_b = resultados['trabajos_grado']
                if tg_b:
                    conn.executemany(
                        '''INSERT INTO trabajos_grado
                           (cedula_director,nombre_director,cedula_estudiante,nombre_estudiante,
                            titulo,programa,año,estado,fecha_sustentacion,fuente)
                           VALUES(?,?,?,?,?,?,?,?,?,?)''',
                        tg_b,
                    )
                    self.progreso.emit(f"Insertados {len(tg_b)} trabajos de grado")

            if resultados.get('libros'):
                _, libros_b = resultados['libros']
                if libros_b:
                    conn.executemany(
                        '''INSERT INTO publicaciones
                           (cedula,titulo,revista_libro,doi_url,issn_isbn,año,
                            tipo,categoria,estado,grupo,fuente)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?)''',
                        libros_b,
                    )
                    self.progreso.emit(f"Insertados {len(libros_b)} libros/capítulos")

            if resultados.get('productos_innovacion'):
                _, pi_b = resultados['productos_innovacion']
                if pi_b:
                    conn.executemany(
                        '''INSERT INTO productos_innovacion
                           (cedula,tipo_producto,nombre,descripcion,año,estado,grupo,fuente)
                           VALUES(?,?,?,?,?,?,?,?)''',
                        pi_b,
                    )
                    self.progreso.emit(f"Insertados {len(pi_b)} productos de innovación")

            conn.commit()

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

    @staticmethod
    def _norm_key(k: str) -> str:
        return re.sub(r'[^a-z0-9]', '_', unidecode(str(k)).lower().strip())

    @classmethod
    def _col(cls, row: dict, *keys) -> str:
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
        df = df.copy()
        df.columns = [cls._norm_key(c) for c in df.columns]
        return df

    @staticmethod
    def _anio(valor):
        if not valor or (isinstance(valor, float) and valor != valor):
            return None
        try:
            return int(str(valor).strip()[:4])
        except (ValueError, TypeError):
            return None

    # ── Integrantes ──

    def _extraer_integrantes(self):
        archivos = [
            'Listado Integrantes Grupos de Investigación UTP  080825.xlsx',
            'Listado Integrantes Grupos de Investigación UTP - 080825.xlsx',
            'Listado Integrantes Grupos de Investigacion UTP  080825.xlsx',
            'Listado Integrantes Grupos de Investigacion UTP - 080825.xlsx',
            'integrantes.xlsx',
        ]
        for archivo in archivos:
            ruta = self.archivos_directorio / "data" / "input" / archivo
            if not ruta.exists():
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

    # ── Extensiones ──

    def _extraer_extensiones(self):
        archivos = [
            'Consolidado Extensión 2024.xlsx',
            'Actividades Extensión enerojulio.xlsx',
            'Actividades Extensión (enero-julio).xlsx',
        ]
        personas_b, ext_b = [], []
        for archivo in archivos:
            ruta = self.archivos_directorio / "data" / "input" / archivo
            if not ruta.exists():
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
            ))
        return personas_b, ext_b

    # ── Producción ──

    def _extraer_produccion(self):
        archivos = [
            'BASE DATOS PRODUCCIÓN 2024.xlsx',
            'BASE DATOS PRODUCCIÓN  2025  CIARP.xlsx',
            'BASE DATOS PRODUCCIÓN  2025 - CIARP.xlsx',
        ]
        personas_b, pub_b = [], []
        for archivo in archivos:
            ruta = self.archivos_directorio / "data" / "input" / archivo
            if not ruta.exists():
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
                ))
        return personas_b, pub_b

    # ── Trabajos de grado ──

    def _extraer_trabajos_grado(self):
        archivos = [
            'Trabajos Grado  Trabajo de Grado.xlsx',
            'TrabajosGrado_TrabajoDeGrado 2024.xlsx',
            'Trabajos Grado - Trabajo de Grado.xlsx',
        ]
        personas_b, tg_b = [], []
        for archivo in archivos:
            ruta = self.archivos_directorio / "data" / "input" / archivo
            if not ruta.exists():
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
                    tg_b.append((
                        cedula_director, director_actual,
                        limpiar_cedula(str(row[2]) if len(row) > 2 and pd.notna(row[2]) else ''),
                        limpiar_texto(str(row[1]) if len(row) > 1 and pd.notna(row[1]) else ''),
                        titulo,
                        limpiar_texto(str(row[3]) if len(row) > 3 and pd.notna(row[3]) else ''),
                        anio,
                        limpiar_texto(str(row[4]) if len(row) > 4 and pd.notna(row[4]) else ''),
                        fecha, archivo,
                    ))
        return personas_b, tg_b

    # ── Libros ──

    def _extraer_libros(self):
        archivos = [
            'Reporte de libros y capítulos publicados.xlsx',
            'Reporte_libros.xlsx',
        ]
        personas_b, pub_b = [], []
        for archivo in archivos:
            ruta = self.archivos_directorio / "data" / "input" / archivo
            if not ruta.exists():
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

    # ── Productos de innovación ──

    def _extraer_productos_innovacion(self):
        ruta = self.archivos_directorio / "data" / "input" / 'info_productos_innovacion.xlsx'
        if not ruta.exists():
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
        ))
        return [], batch

    # ── Proyectos ──

    def cargar_proyectos(self):
        self.progreso.emit("Cargando proyectos de investigación...")
        archivos_posibles = [
            "Proyectos de investigación registrados en 2024.xlsx",
            "Proyectos de investigacion registrados en 2024.xlsx",
            "proyectos_investigacion_2024.xlsx",
            "proyectos 2024.xlsx"
        ]
        archivo_encontrado = None
        for archivo in archivos_posibles:
            ruta = self.archivos_directorio / "data" / "input" / archivo
            if not ruta.exists():
                ruta = self.archivos_directorio / archivo
            if ruta.exists():
                archivo_encontrado = archivo
                break
        if not archivo_encontrado:
            self.progreso.emit("⚠ No se encontró archivo de proyectos de investigación")
            return
        ruta = self.archivos_directorio / "data" / "input" / archivo_encontrado
        if not ruta.exists():
            ruta = self.archivos_directorio / archivo_encontrado
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", UserWarning)
                df_raw = pd.read_excel(ruta, sheet_name=0, header=None, dtype=str, engine="openpyxl")
            df_raw = df_raw.fillna("")
            hdr_row = self._find_header_row_proyectos(df_raw)
            if hdr_row is None:
                self.progreso.emit(f"⚠ No se detectó encabezado en {archivo_encontrado}. Intentando lectura estándar...")
                df = pd.read_excel(ruta, engine='openpyxl')
            else:
                self.progreso.emit(f"✓ Encabezado detectado en fila {hdr_row+1}")
                colmap = self._pick_columns_proyectos(df_raw.iloc[hdr_row, :])
                self.progreso.emit(f"Columnas detectadas: {sum(1 for v in colmap.values() if v is not None)}/{len(colmap)}")
                df = pd.DataFrame()
                for col_name, col_idx in colmap.items():
                    if col_idx is not None:
                        df[col_name] = df_raw.iloc[hdr_row+1:, col_idx].astype(str).str.strip()
            count = 0
            for _, row in df.iterrows():
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
                if pd.notna(año) and año:
                    try:
                        año = int(str(año).split('.')[0])
                    except:
                        año = None
                cedula_principal = self.db.obtener_cedula_principal(cedula)
                cursor = self.db.conn.cursor()
                cursor.execute('''
                    INSERT INTO proyectos
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

    def _find_header_row_proyectos(self, df, max_scan=50):
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

    def _pick_columns_proyectos(self, header):
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

    # ── Propiedad intelectual ──

    def cargar_propiedad_intelectual(self):
        self.progreso.emit("Cargando registros de propiedad intelectual...")
        archivos_posibles = [
            "CGT0104  No de productos resultados de investigacion 31072025.xlsx",
            "CGT0104  No de productos resultados de investigacion 31122024.xlsx"
        ]
        count = 0
        for archivo in archivos_posibles:
            ruta = self.archivos_directorio / "data" / "input" / archivo
            if not ruta.exists():
                ruta = self.archivos_directorio / archivo
            if not ruta.exists():
                continue
            try:
                wb = pd.ExcelFile(ruta, engine='openpyxl')
                for sheet_name in wb.sheet_names:
                    if 'Soporte' in sheet_name or 'Reg Propiedad' in sheet_name:
                        df = pd.read_excel(ruta, sheet_name=sheet_name, engine='openpyxl')
                        for _, row in df.iterrows():
                            responsable = limpiar_texto(row.get('Responsables') or row.get('Responsable') or '')
                            tipo_producto = limpiar_texto(row.get('Tipo de producto') or '')
                            nombre = limpiar_texto(row.get('Nombre del producto') or row.get('Nombre') or '')
                            if not responsable and not nombre:
                                continue
                            cursor = self.db.conn.cursor()
                            cursor.execute('''
                                INSERT INTO propiedad_intelectual
                                (responsable, tipo_producto, tipo_patente, nombre_producto,
                                 numero_registro, proyecto, fecha_aprobacion, entidad, facultad, fuente)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                                f"{archivo} :: {sheet_name}"
                            ))
                            count += 1
                self.db.conn.commit()
            except Exception as e:
                self.progreso.emit(f"Error en {archivo}: {str(e)}")
        if count > 0:
            self.progreso.emit(f"Cargados {count} registros de propiedad intelectual")
