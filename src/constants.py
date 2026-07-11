"""
Constantes globales para el sistema ReportSoft - Consolidados.

Fuente autoritativa de:
  - CLASIF_957: clasificación oficial Convocatoria 957 MinCiencias (fusión de ambas vistas)
  - VARIANTES_NOMBRES: alias de nombres de hojas GrupLAC
  - VENTANAS_957: ventanas de evaluación y pesos (lambda) por tipo de producto
  - MAPEO_CATEGORIA_PUBLICACION: mapeo de categoría de artículo → indicador 957
  - Constantes de color y orden para la interfaz
"""

from collections import OrderedDict

# =============================================================================
# DICCIONARIO OFICIAL CONVOCATORIA 957 MINCIENCIAS
# Clave: nombre oficial del producto (como aparece en GrupLAC)
# Valor: (categoría_principal, subcategoría)
# =============================================================================
CLASIF_957 = OrderedDict([
    # ------------------------------------------------------------------
    # 1. Generación de nuevo conocimiento
    # ------------------------------------------------------------------
    ("Artículos publicados", (
        "1. Generación de nuevo conocimiento",
        "Artículos de investigación (A1, A2, B, C y D)",
    )),
    ("Notas científicas", (
        "1. Generación de nuevo conocimiento",
        "Notas científicas (A1, A2, B, C y D)",
    )),
    ("Libros publicados", (
        "1. Generación de nuevo conocimiento",
        "Libros resultado de investigación",
    )),
    ("Capítulos de libro publicados", (
        "1. Generación de nuevo conocimiento",
        "Capítulos en libro resultado de investigación",
    )),
    ("Libros de formación", (
        "1. Generación de nuevo conocimiento",
        "Libros de formación (Q1)",
    )),
    ("Nuevas variedades vegetal", (
        "1. Generación de nuevo conocimiento",
        "Innovación biológica",
    )),
    ("Nuevas variedades animal", (
        "1. Generación de nuevo conocimiento",
        "Innovación biológica",
    )),
    ("Poblaciones mejoradas de razas pecuarias", (
        "1. Generación de nuevo conocimiento",
        "Innovación biológica",
    )),
    ("Producción en arte, arquitectura y diseño", (
        "1. Generación de nuevo conocimiento",
        "Investigación–creación",
    )),
    ("Obras o productos", (
        "1. Generación de nuevo conocimiento",
        "Investigación–creación",
    )),

    # ------------------------------------------------------------------
    # 2. Desarrollo tecnológico e innovación
    # ------------------------------------------------------------------
    ("Diseños industriales", (
        "2. Desarrollo tecnológico e innovación",
        "Producto tecnológico",
    )),
    ("Esquemas de trazados de circuito integrado", (
        "2. Desarrollo tecnológico e innovación",
        "Producto tecnológico",
    )),
    ("Softwares", (
        "2. Desarrollo tecnológico e innovación",
        "Producto tecnológico",
    )),
    ("Plantas piloto", (
        "2. Desarrollo tecnológico e innovación",
        "Producto tecnológico",
    )),
    ("Prototipos", (
        "2. Desarrollo tecnológico e innovación",
        "Producto tecnológico",
    )),
    ("Signos distintivos", (
        "2. Desarrollo tecnológico e innovación",
        "Producto tecnológico",
    )),
    ("Productos nutracéuticos", (
        "2. Desarrollo tecnológico e innovación",
        "Producto tecnológico",
    )),
    ("Nuevos registros científicos", (
        "2. Desarrollo tecnológico e innovación",
        "Certificación y validación",
    )),
    ("Innovaciones en Procesos y Procedimientos", (
        "2. Desarrollo tecnológico e innovación",
        "Innovación organizacional",
    )),
    ("Innovaciones generadas en la Gestión Empresaria", (
        "2. Desarrollo tecnológico e innovación",
        "Innovación organizacional",
    )),
    ("Empresas de base tecnológica", (
        "2. Desarrollo tecnológico e innovación",
        "Innovación empresarial",
    )),
    ("Industrias creativas y culturales", (
        "2. Desarrollo tecnológico e innovación",
        "Innovación empresarial",
    )),
    ("Regulaciones y Normas", (
        "2. Desarrollo tecnológico e innovación",
        "Marco normativo",
    )),
    ("Reglamentos técnicos", (
        "2. Desarrollo tecnológico e innovación",
        "Marco normativo",
    )),
    ("Guías de práctica clínica", (
        "2. Desarrollo tecnológico e innovación",
        "Marco normativo",
    )),
    ("Protocolos de vigilancia epidemiológica", (
        "2. Desarrollo tecnológico e innovación",
        "Marco normativo",
    )),
    ("Proyectos de ley", (
        "2. Desarrollo tecnológico e innovación",
        "Legislación",
    )),
    ("Conceptos técnicos", (
        "2. Desarrollo tecnológico e innovación",
        "Transferencia tecnológica",
    )),
    ("Cartas, mapas o similares", (
        "2. Desarrollo tecnológico e innovación",
        "Otros productos tecnológicos",
    )),
    ("Informes técnicos", (
        "2. Desarrollo tecnológico e innovación",
        "Otros productos tecnológicos",
    )),
    ("Otros productos tecnológicos", (
        "2. Desarrollo tecnológico e innovación",
        "Otros productos tecnológicos",
    )),

    # ------------------------------------------------------------------
    # 3. Apropiación social del conocimiento
    # ------------------------------------------------------------------
    (
        "Procesos de apropiación social del Conocimiento para el "
        "fortalecimiento o solución de asuntos de interés social",
        ("3. Apropiación social del conocimiento", "Interacción social"),
    ),
    ("Estrategias Pedagógicas para el fomento de la CTeI", (
        "3. Apropiación social del conocimiento",
        "Interacción social",
    )),
    (
        "Proceso de Apropiación Social del Conocimiento para la generación "
        "de insumos de política pública y normatividad",
        ("3. Apropiación social del conocimiento", "Política pública"),
    ),
    (
        "Proceso de apropiación social del Conocimiento para el "
        "fortalecimiento de cadenas productivas",
        ("3. Apropiación social del conocimiento", "Desarrollo productivo"),
    ),
    (
        "Productos de apropiación social del conocimiento resultado del "
        "trabajo conjunto entre un Centro de Ciencia y un grupo de investigación",
        ("3. Apropiación social del conocimiento", "Articulación institucional"),
    ),
    ("Espacios de Participación Ciudadana", (
        "3. Apropiación social del conocimiento",
        "Participación ciudadana",
    )),
    ("Participación Ciudadana en Proyectos de CTI", (
        "3. Apropiación social del conocimiento",
        "Participación ciudadana",
    )),

    # ------------------------------------------------------------------
    # 4. Divulgación pública de la ciencia
    # ------------------------------------------------------------------
    ("Eventos Científicos", (
        "4. Divulgación pública de la ciencia",
        "Circulación especializada",
    )),
    ("Redes de Conocimiento Especializado", (
        "4. Divulgación pública de la ciencia",
        "Circulación especializada",
    )),
    ("Talleres de Creación", (
        "4. Divulgación pública de la ciencia",
        "Circulación especializada",
    )),
    ("Eventos Artísticos", (
        "4. Divulgación pública de la ciencia",
        "Circulación especializada",
    )),
    ("Documentos de trabajo", (
        "4. Divulgación pública de la ciencia",
        "Producción técnica",
    )),
    ("Informes de investigación", (
        "4. Divulgación pública de la ciencia",
        "Producción técnica",
    )),
    ("Consultorías científico-tecnológicas", (
        "4. Divulgación pública de la ciencia",
        "Producción técnica",
    )),
    ("Ediciones", (
        "4. Divulgación pública de la ciencia",
        "Producción editorial",
    )),
    ("Nuevas secuencias genéticas", (
        "4. Divulgación pública de la ciencia",
        "Producción editorial",
    )),
    ("Publicaciones editoriales no especializadas", (
        "4. Divulgación pública de la ciencia",
        "Divulgación multiformato",
    )),
    ("Producciones de contenido digital - Audiovisual", (
        "4. Divulgación pública de la ciencia",
        "Divulgación multiformato",
    )),
    ("Producciones de contenido digital - Sonoro", (
        "4. Divulgación pública de la ciencia",
        "Divulgación multiformato",
    )),
    ("Producciones de contenido digital - Recursos gráficos", (
        "4. Divulgación pública de la ciencia",
        "Divulgación multiformato",
    )),
    (
        "Divulgación Pública de la Ciencia producción de estrategias "
        "y contenidos transmedia",
        ("4. Divulgación pública de la ciencia", "Divulgación multiformato"),
    ),
    ("Desarrollo web", (
        "4. Divulgación pública de la ciencia",
        "Divulgación multiformato",
    )),
    ("Generación de Contenido Virtual", (
        "4. Divulgación pública de la ciencia",
        "Divulgación multiformato",
    )),
    ("Generación de Contenido Multimedia", (
        "4. Divulgación pública de la ciencia",
        "Divulgación multiformato",
    )),
    ("Generación de Contenido Impreso", (
        "4. Divulgación pública de la ciencia",
        "Divulgación multiformato",
    )),
    ("Generaciones de contenido de audio", (
        "4. Divulgación pública de la ciencia",
        "Divulgación multiformato",
    )),
    ("Estrategias de Comunicación del Conocimiento", (
        "4. Divulgación pública de la ciencia",
        "Divulgación multiformato",
    )),
    ("Libros de divulgación y/o Compilación de divulgación", (
        "4. Divulgación pública de la ciencia",
        "Producción bibliográfica",
    )),
    ("Otra publicación divulgativa", (
        "4. Divulgación pública de la ciencia",
        "Producción bibliográfica",
    )),
    ("Manuales y Guías Especializadas", (
        "4. Divulgación pública de la ciencia",
        "Producción bibliográfica",
    )),
    ("Otros artículos publicados", (
        "4. Divulgación pública de la ciencia",
        "Producción bibliográfica",
    )),
    ("Traducciones", (
        "4. Divulgación pública de la ciencia",
        "Producción bibliográfica",
    )),

    # ------------------------------------------------------------------
    # 5. Formación de recurso humano en CTeI
    # ------------------------------------------------------------------
    ("Programa académico de doctorado", (
        "5. Formación de recurso humano en CTeI",
        "Formación académica",
    )),
    ("Programa académico de maestría", (
        "5. Formación de recurso humano en CTeI",
        "Formación académica",
    )),
    ("Otro programa académico", (
        "5. Formación de recurso humano en CTeI",
        "Formación académica",
    )),
    ("Curso de doctorado", (
        "5. Formación de recurso humano en CTeI",
        "Cursos",
    )),
    ("Curso de maestría", (
        "5. Formación de recurso humano en CTeI",
        "Cursos",
    )),
    ("Curso especializado de extensión", (
        "5. Formación de recurso humano en CTeI",
        "Extensión",
    )),
    ("Curso de Corta Duración Dictados", (
        "5. Formación de recurso humano en CTeI",
        "Extensión",
    )),
    ("Trabajos dirigidos/tutorías", (
        "5. Formación de recurso humano en CTeI",
        "Dirección académica",
    )),
    ("Proyectos", (
        "5. Formación de recurso humano en CTeI",
        "Proyectos de investigación",
    )),
    ("Asesorías al Programa Ondas", (
        "5. Formación de recurso humano en CTeI",
        "Formación temprana",
    )),
    ("Jurado/Comisiones evaluadoras de trabajo de grado", (
        "5. Formación de recurso humano en CTeI",
        "Evaluador",
    )),
    ("Participación en comités de evaluación", (
        "5. Formación de recurso humano en CTeI",
        "Evaluador",
    )),
    ("Demás trabajos", (
        "5. Formación de recurso humano en CTeI",
        "Otros",
    )),
])

CATEGORIAS_PRINCIPALES = [
    "1. Generación de nuevo conocimiento",
    "2. Desarrollo tecnológico e innovación",
    "3. Apropiación social del conocimiento",
    "4. Divulgación pública de la ciencia",
    "5. Formación de recurso humano en CTeI",
]

# =============================================================================
# VARIANTES DE NOMBRES (hojas GrupLAC → nombre oficial en CLASIF_957)
# Clave: nombre normalizado (sin tildes, minúsculas, sin espacios dobles)
# Valor: nombre oficial tal como aparece en CLASIF_957
# =============================================================================
VARIANTES_NOMBRES = {
    # Artículos
    "articulos publicados": "Artículos publicados",
    "notas cientificas": "Notas científicas",
    # Libros / capítulos
    "capitulos de libro publicados": "Capítulos de libro publicados",
    "otros libros publicados": "Otros artículos publicados",
    # Trabajos dirigidos — alias de nombre con tipografía errada
    "trabajos dirigidos/tutorias": "Trabajos dirigidos/tutorías",
    "trabajos dirigidos/turorias": "Trabajos dirigidos/tutorías",
    "trabajos dirigidostutorias": "Trabajos dirigidos/tutorías",
    "trabajos dirigidosturorias": "Trabajos dirigidos/tutorías",
    # Guías
    "guias de practica clinica": "Guías de práctica clínica",
    # Consultorías
    "consultorias cientifico-tecnologicas": "Consultorías científico-tecnológicas",
    # Software
    "software": "Softwares",
    # Eventos
    "eventos cientificos": "Eventos Científicos",
    # Jurado
    "juradocomisiones evaluadoras": "Jurado/Comisiones evaluadoras de trabajo de grado",
    "jurado comisiones evaluadoras": "Jurado/Comisiones evaluadoras de trabajo de grado",
    # Libros divulgación (nombres truncados por GrupLAC)
    "libros de divulgacion yo comp": "Libros de divulgación y/o Compilación de divulgación",
    "libros de divulgacion y/o comp": "Libros de divulgación y/o Compilación de divulgación",
    # Contenido digital (nombres truncados)
    "generacion de contenido virtua": "Generación de Contenido Virtual",
    "generacion de contenido virtual": "Generación de Contenido Virtual",
    "generacion de contenido multim": "Generación de Contenido Multimedia",
    "generacion de contenido multimedia": "Generación de Contenido Multimedia",
    "generacion de contenido impres": "Generación de Contenido Impreso",
    "generacion de contenido impreso": "Generación de Contenido Impreso",
    "generaciones de contenido de a": "Generaciones de contenido de audio",
    "generaciones de contenido de audio": "Generaciones de contenido de audio",
    # Estrategias (nombres truncados)
    "estrategias pedagogicas para e": "Estrategias Pedagógicas para el fomento de la CTeI",
    "estrategias pedagogicas para el fomento": "Estrategias Pedagógicas para el fomento de la CTeI",
    "estrategias de comunicacion de": "Estrategias de Comunicación del Conocimiento",
    "estrategias de comunicacion del conocimiento": "Estrategias de Comunicación del Conocimiento",
    # Otras publicaciones
    "otra publicacion divulgativa": "Otra publicación divulgativa",
    "traducciones filologicas y edi": "Traducciones",
    "traducciones filologicas": "Traducciones",
}

# Hojas de GrupLAC que NO son productos (se ignoran en clasificación)
HOJAS_NO_PRODUCTO = {
    "datos basicos",
    "instituciones",
    "integrantes del grupo",
    "lineas de investigacion",
    "plan estrategico",
}

# =============================================================================
# COLORES PARA INTERFAZ
# =============================================================================
COLORES_CATEGORIA = {
    "1. Generación de nuevo conocimiento": "#2E86AB",
    "2. Desarrollo tecnológico e innovación": "#A23B72",
    "3. Apropiación social del conocimiento": "#F18F01",
    "4. Divulgación pública de la ciencia": "#C73E1D",
    "5. Formación de recurso humano en CTeI": "#3B8C66",
    "Sin clasificar": "#7F8C8D",
}

COLORES_CATEGORIA_HEX = {
    "1. Generación de nuevo conocimiento": "D6EAF8",
    "2. Desarrollo tecnológico e innovación": "F5EEF8",
    "3. Apropiación social del conocimiento": "FEF5E7",
    "4. Divulgación pública de la ciencia": "FADBD8",
    "5. Formación de recurso humano en CTeI": "D5F5E3",
    "Sin clasificar": "E5E7E9",
}

# =============================================================================
# VENTANAS DE EVALUACIÓN 957 Y PESOS (LAMBDA)
# Basado en la metodología oficial de MinCiencias Convocatoria 957.
# ventana: años hacia atrás desde el año base de la convocatoria.
# lambda_val: peso del producto en el cálculo del indicador.
# indicador: a qué indicador del modelo 957 contribuye.
#
# NOTA: Verificar contra el documento oficial de la convocatoria vigente.
# =============================================================================
VENTANAS_957 = {
    # Categoría publicación interna → (ventana_años, lambda, indicador)
    "A1": {"ventana": 5, "lambda": 1.00, "indicador": "TOP"},
    "A2": {"ventana": 5, "lambda": 0.75, "indicador": "TIPO_A"},
    "B":  {"ventana": 5, "lambda": 0.50, "indicador": "TIPO_B"},
    "C":  {"ventana": 5, "lambda": 0.25, "indicador": "AP"},
    "D":  {"ventana": 5, "lambda": 0.10, "indicador": "AP"},
    # SJR quartiles (sinónimos)
    "Q1": {"ventana": 5, "lambda": 1.00, "indicador": "TOP"},
    "Q2": {"ventana": 5, "lambda": 0.75, "indicador": "TIPO_A"},
    "Q3": {"ventana": 5, "lambda": 0.50, "indicador": "TIPO_B"},
    "Q4": {"ventana": 5, "lambda": 0.25, "indicador": "AP"},
    # Libros
    "LIBRO": {"ventana": 10, "lambda": 2.00, "indicador": "TOP"},
    "CAPITULO": {"ventana": 5, "lambda": 1.00, "indicador": "TIPO_A"},
    # Formación de recurso humano
    "DOCTORADO": {"ventana": 5, "lambda": 0.50, "indicador": "DPC"},
    "MAESTRIA":  {"ventana": 3, "lambda": 0.25, "indicador": "FR_A"},
    "ESPECIALIZACION": {"ventana": 3, "lambda": 0.10, "indicador": "FR_B"},
    "PREGRADO":  {"ventana": 3, "lambda": 0.05, "indicador": "FR_B"},
    # Innovación / propiedad intelectual
    "PATENTE":   {"ventana": 10, "lambda": 1.00, "indicador": "TOP"},
    "SOFTWARE":  {"ventana": 5,  "lambda": 0.50, "indicador": "TIPO_B"},
    "PROTOTIPO": {"ventana": 5,  "lambda": 0.50, "indicador": "TIPO_B"},
}

# =============================================================================
# INDICADORES DEL MODELO 957
# =============================================================================
INDICADORES_957 = ["TOP", "TIPO_A", "TIPO_B", "AP", "DPC", "FR_A", "FR_B"]

ORDEN_CATEGORIAS_MINCIENCIAS = ["D", "C", "B", "A", "A1"]

# =============================================================================
# UMBRALES POR ÁREA DE CONOCIMIENTO (medicion_957.xlsx → hoja "cuartiles")
# Para cada indicador del modelo 957, su etiqueta equivalente en la hoja
# "cuartiles" del documento oficial de medición Conv. 957.
# =============================================================================
INDICADOR_957_A_CUARTIL = {
    "TOP":    "Nuevo Conocimiento TOP",
    "TIPO_A": "Nuevo Conocimiento A",
    "TIPO_B": "Nuevo Conocimiento B",
    "AP":     "Apropiación Social y del Conocimiento",
    "DPC":    "Divulgación Pública de la Ciencia",
    "FR_A":   "Formación de Recurso Humano A",
    "FR_B":   "Formación de Recurso Humano B",
}

# Categoría objetivo (a la que se quiere ascender) → cuartil mínimo requerido,
# según columna de la hoja "cuartiles" (min/q4/q3/q2/max).
# Validado empíricamente: A1 ≈ Q1 (>=q2), A ≈ Q2 (>=q3), B ≈ Q3 (>=q4), C ≈ Q4 (>=min).
CUARTIL_OBJETIVO_POR_CATEGORIA = {
    "C":  "min",
    "B":  "q4",
    "A":  "q3",
    "A1": "q2",
}

# =============================================================================
# CÓDIGOS DE LA HOJA "indicadores"/"productos" DE medicion_957.xlsx
# → indicador del modelo 957 (INDICADORES_957).
# Estos valores ya vienen calculados y validados por MinCiencias a partir de
# los productos del documento oficial de medición ("carpeta pdf").
# =============================================================================
CODIGO_957_A_INDICADOR = {
    "NC_TOP": "TOP",
    "NC_A":   "TIPO_A",
    "NC_B":   "TIPO_B",
    "ASC":    "AP",
    "DPC":    "DPC",
    "FRH_A":  "FR_A",
    "FRH_B":  "FR_B",
}

# =============================================================================
# ARCHIVOS FUENTE ESPERADOS (pestaña Inicio + CargadorDatosIntegrado)
# Cada categoría representa un dato de entrada que el usuario debe mantener
# actualizado en la carpeta base del programa.
#   clave: identificador interno
#   label: nombre mostrado al usuario en la pestaña Inicio
#   variantes: nombres de archivo reconocidos (por inconsistencias históricas
#       de tildes/guiones); el cargador busca el primero que exista
#   canónico (variantes[0]): nombre con el que se guarda el archivo que el
#       usuario carga desde la pestaña Inicio; es el mismo nombre que ya
#       vigila CargadorDatosIntegrado para detectar cambios por fecha.
# =============================================================================
ARCHIVOS_FUENTE_957 = [
    {
        "clave": "integrantes",
        "label": "Integrantes de Grupos de Investigación",
        "variantes": [
            "Listado Integrantes Grupos de Investigación UTP  080825.xlsx",
            "Listado Integrantes Grupos de Investigación UTP - 080825.xlsx",
            "Listado Integrantes Grupos de Investigacion UTP  080825.xlsx",
            "Listado Integrantes Grupos de Investigacion UTP - 080825.xlsx",
        ],
    },
    {
        "clave": "extension",
        "label": "Actividades de Extensión",
        "variantes": [
            "Consolidado Extensión 2024.xlsx",
            "Actividades Extensión enerojulio.xlsx",
            "Actividades Extensión (enero-julio).xlsx",
        ],
    },
    {
        "clave": "produccion_2024",
        "label": "Producción 2024 (publicaciones)",
        "variantes": [
            "BASE DATOS PRODUCCIÓN 2024.xlsx",
        ],
    },
    {
        "clave": "produccion_2025_ciarp",
        "label": "Producción 2025 — CIARP (publicaciones)",
        "variantes": [
            "BASE DATOS PRODUCCIÓN  2025  CIARP.xlsx",
            "BASE DATOS PRODUCCIÓN  2025 - CIARP.xlsx",
        ],
    },
    {
        "clave": "trabajos_grado",
        "label": "Trabajos de Grado",
        "variantes": [
            "Trabajos Grado  Trabajo de Grado.xlsx",
            "TrabajosGrado_TrabajoDeGrado 2024.xlsx",
            "Trabajos Grado - Trabajo de Grado.xlsx",
        ],
    },
    {
        "clave": "libros",
        "label": "Libros y Capítulos Publicados",
        "variantes": [
            "Reporte_libros.xlsx",
            "Reporte de libros y capítulos publicados.xlsx",
        ],
    },
    {
        "clave": "innovacion",
        "label": "Productos de Innovación",
        "variantes": [
            "info_productos_innovacion.xlsx",
        ],
    },
    {
        "clave": "proyectos",
        "label": "Proyectos de Investigación Registrados",
        "variantes": [
            "Proyectos de investigación registrados en 2024.xlsx",
            "Proyectos de investigacion registrados en 2024.xlsx",
            "proyectos_investigacion_2024.xlsx",
            "proyectos 2024.xlsx",
        ],
    },
    {
        "clave": "cgt0104_2025",
        "label": "CGT0104 — Corte 31/07/2025 (productos de investigación)",
        "variantes": [
            "CGT0104 - No de productos resultados de investigacion 31072025.xlsx",
            "CGT0104  No de productos resultados de investigacion 31072025.xlsx",
        ],
    },
    {
        "clave": "cgt0104_2024",
        "label": "CGT0104 — Corte 31/12/2024 (productos de investigación)",
        "variantes": [
            "CGT0104 - No de productos resultados de investigacion 31122024.xlsx",
            "CGT0104  No de productos resultados de investigacion 31122024.xlsx",
        ],
    },
]

# =============================================================================
# COLUMNAS CONOCIDAS POR CATEGORÍA (detección de columnas nuevas)
# Nombres ya normalizados con utils.normalizar_columna (minúsculas, sin tildes,
# no-alfanumérico → "_"). Son las columnas que los extractores de
# CargadorDatosIntegrado efectivamente leen para esa categoría.
#
# Solo se incluyen categorías con encabezado de columnas estable (hoja tabular
# con nombres de columna fijos). "trabajos_grado", "libros" y "proyectos"
# quedan fuera porque su lectura es posicional o usa detección dinámica de
# encabezado, así que no tiene sentido comparar "columnas nuevas".
#
# Cuando un Excel cargado trae columnas que no están en este set, la pestaña
# Inicio le pregunta al usuario si desea continuar; si acepta, esas columnas
# se guardan completas (como JSON) en el campo "datos_adicionales" de la fila
# (para las categorías cuya tabla tiene esa columna; "integrantes" no la
# tiene porque alimenta "personas"/"grupos", compartidas con otras fuentes,
# pero igual se valida que el archivo traiga el formato esperado).
# =============================================================================
COLUMNAS_CONOCIDAS_POR_CATEGORIA = {
    "integrantes": {
        "numero_documento", "cedula", "nombres", "nombre",
        "nombre_grupo", "grupo", "facultad", "email", "tipo",
    },
    "extension": {
        "cedula", "nombre_responsable", "facultad_dependencia", "facultad",
        "fecha_inicial", "grupo_semillero_de_investigacion", "grupo",
        "nombre_actividad", "tipo", "modalidad", "estado", "fecha_final",
        "poblacion_beneficiaria", "financiacion_interna",
        "fuente_financiacion_externa",
    },
    "produccion_2024": {
        "cedula", "autores", "autor", "nombre", "dependencia", "facultad",
        "nombre_del_trabajo", "titulo", "revista_o_libro", "revista_libro",
        "doi_url", "doi", "issn_isbn", "issn", "ano_de_la_publicacion", "ano",
        "tipo", "categoria", "estado", "grupo",
    },
    "innovacion": {
        "nombre", "titulo", "producto", "ano_de_registro", "ano",
        "tipo_de_producto", "tipo", "descripcion", "estado",
        "grupo_de_investigacion", "grupo",
    },
    "cgt0104_2025": {
        "responsables", "responsable", "tipo_de_producto", "nombre_del_producto",
        "nombre", "tipo_de_patente", "no_de_registro", "proyecto_de_investigacion",
        "fecha_de_aprobacion", "entidad_que_lo_expide", "facultad",
    },
}
# La categoría "produccion_2025_ciarp" comparte el mismo extractor/columnas que
# "produccion_2024", y "cgt0104_2024" comparte las de "cgt0104_2025".
COLUMNAS_CONOCIDAS_POR_CATEGORIA["produccion_2025_ciarp"] = COLUMNAS_CONOCIDAS_POR_CATEGORIA["produccion_2024"]
COLUMNAS_CONOCIDAS_POR_CATEGORIA["cgt0104_2024"] = COLUMNAS_CONOCIDAS_POR_CATEGORIA["cgt0104_2025"]

# =============================================================================
# COLUMNAS "CLAVE" (discriminantes) POR CATEGORÍA
# Subconjunto de COLUMNAS_CONOCIDAS_POR_CATEGORIA que de verdad distingue ese
# tipo de archivo de cualquier otro. Columnas genéricas como "nombre", "email"
# o "cedula" aparecen en casi cualquier planilla institucional, así que por sí
# solas NO bastan para considerar que un archivo "sí coincide": un archivo de
# notas también trae "nombre" y "email", por ejemplo. La pestaña Inicio exige
# al menos una columna clave presente para no mostrar la advertencia fuerte de
# "el archivo no coincide con el formato esperado".
# =============================================================================
COLUMNAS_CLAVE_POR_CATEGORIA = {
    "integrantes": {"numero_documento", "cedula", "nombre_grupo", "grupo"},
    "extension": {
        "nombre_actividad", "fecha_inicial", "modalidad",
        "grupo_semillero_de_investigacion", "poblacion_beneficiaria",
    },
    "produccion_2024": {
        "nombre_del_trabajo", "revista_o_libro", "revista_libro",
        "doi_url", "doi", "issn_isbn", "issn", "ano_de_la_publicacion",
    },
    "innovacion": {"tipo_de_producto", "ano_de_registro", "grupo_de_investigacion"},
    "cgt0104_2025": {
        "tipo_de_producto", "no_de_registro", "proyecto_de_investigacion",
        "fecha_de_aprobacion", "entidad_que_lo_expide",
    },
}
COLUMNAS_CLAVE_POR_CATEGORIA["produccion_2025_ciarp"] = COLUMNAS_CLAVE_POR_CATEGORIA["produccion_2024"]
COLUMNAS_CLAVE_POR_CATEGORIA["cgt0104_2024"] = COLUMNAS_CLAVE_POR_CATEGORIA["cgt0104_2025"]

# Sección del modelo 957 (hoja "productos"/"indicadores" de medicion_957.xlsx)
# a la que aporta cada indicador. FR_A/FR_B (formación de recurso humano) no
# tienen desglose de productos en el documento oficial.
SECCION_957_POR_INDICADOR = {
    "TOP":    "NC_TOP",
    "TIPO_A": "NC_A",
    "TIPO_B": "NC_B",
    "AP":     "ASC",
    "DPC":    "DPC",
}