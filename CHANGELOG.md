# Changelog

Formato libre, en español, más reciente primero. Para el detalle técnico de
los bugs del pipeline de verificación ver
[`docs/VERIFICACION_GRUPLAC.md`](docs/VERIFICACION_GRUPLAC.md) — acá solo el
resumen de qué cambió y por qué.

## 2026-07-14 a 2026-07-17

### Agregado
- **`UI_clasificacion/`** (nuevo, módulo separado de la app principal, ver
  su propio [`README.md`](UI_clasificacion/README.md)): categoría 957
  vigente de cada grupo (100% desde el scraping GrupLAC, siempre
  actualizada) + una proyección ESTIMADA de cómo quedaría cada grupo si la
  convocatoria se recalculara hoy, usando el motor ya validado
  `Simulador957` alimentado con producción de artículos posterior al cierre
  de la ventana oficial (2023), clasificados por ISSN contra el dataset
  abierto de Publindex (`datos.gov.co`, id `mwmn-inyg`). Panel de detalle
  por grupo con el cálculo metodológico completo (índices, cuartiles del
  área, condiciones oficiales A1/A/B/C, trazabilidad artículo por artículo).
- **`src/chatbot_investigacion.py`** (nuevo): chatbot analítico con
  tool-calling real contra Ollama (14 herramientas: consultas de grupos/
  investigadores/producción + 3 herramientas de gráficas + cumplimiento/
  faltantes GrupLAC + cambio de pestaña). Corrige un bug real del prototipo
  previo (`vista_chat_ollama.py.bak`, nunca integrado): no cerraba el ciclo
  de tool-calling (ejecutaba la herramienta pero nunca le devolvía el
  resultado al modelo para la respuesta final). Probado end-to-end con
  Ollama + `qwen2.5:7b` real. **Pestaña desactivada por ahora** (a pedido
  del usuario) — código intacto, ver comentarios en
  `main_10.py::VentanaPrincipal.setup_ui`.
- **Prácticas → Trabajos de Grado**: nuevo extractor
  (`CargadorDatosIntegrado._extraer_practicas`) para el reporte de
  prácticas de la universidad, con distinción CONDUCENTE/NO CONDUCENTE
  (columna `calificacion` de `trabajos_grado`) — solo lo conducente cuenta
  para GrupLAC (Reportes por Grupo, Seguimiento Grupos, simulador interno);
  lo no conducente solo se muestra en Búsqueda Persona con un badge.
- **Acumulación de fuentes** ("Agregar archivo" en Inicio): agregar un
  archivo nuevo ya no reemplaza el anterior de la misma categoría, se suman
  ambos (`DatabaseManager.obtener_fuentes_adicionales`/
  `registrar_fuente_adicional`).
- Selector de rango de años y categoría "Propiedad Intelectual" en Reportes
  por Grupo (`VistaGrupos`).
- **Procedencia (Grupo / Semillero / Ambos)** en Búsqueda de integrantes:
  columna nueva `grupos.origen` (migración) + mapeo semillero → grupo
  adscrito desde `data/Reporte Semilleros con Grupo adscrito.xlsx`
  (`CargadorDatosIntegrado._grupo_adscrito_semilleros`) — sin esto, ~44
  personas que solo aparecían en un semillero (GrupLAC no rastrea
  semilleros como entidad propia) quedaban sin poder atribuirse a ningún
  grupo verificable.
- "Guardar como" en todos los botones de exportación (antes guardaban en
  una ruta fija sin preguntar).

### Corregido
- **Bug real de atribución de grupo** en la verificación GrupLAC
  (`comparador_gruplac_scrapeado.construir_df_interno`): 67/263 proyectos
  (25%) tenían un campo `grupo` de texto libre que no coincidía con la
  membresía real del responsable — el faltante quedaba invisible en el
  panel de Cumplimiento del grupo real. Ahora se emite también para los
  grupos reales del responsable, no solo el de texto libre.
- **Falso positivo por título genérico** (ej. "ESTUDIANTE EN PRACTICA",
  compartido por 103 estudiantes distintos): se cruza el nombre del
  estudiante (scrapeado del bloque GrupLAC) contra el interno antes de
  confirmar a ciegas un match de título repetido.
- **Pérdida silenciosa de filas por columnas faltantes**
  (`_serie_o_vacia` en `main_10.py`, portado desde `loader.py`): un
  `pd.Series(dtype=str)` suelto como valor por defecto tiene índice vacío;
  usado dentro de un `zip()` junto a columnas con datos, corta TODO el
  resultado a longitud 0 sin ningún error visible. Causaba que la base
  CIARP cargara **0 publicaciones** en vez de 150 (le faltaba la columna
  `doi_url`). Corregido en los 4 extractores afectados (integrantes,
  extensiones, producción, innovación).
- **Miscategorización de archivo nuevo** ("Informe Extensión..." se
  clasificó como "integrantes" en vez de "extensión" porque el detector de
  categoría solo miraba una hoja llamada "Consolidado" y el archivo nuevo
  usa "Datos"): se agregó un extractor dedicado
  (`_extraer_extension_informe`) y se mejoró la detección de categoría en
  Inicio para que gane automáticamente la de más columnas coincidentes en
  vez de exigir coincidencia única (menos popups de "tipo ambiguo/no
  reconocido").
- **Pérdida de coautores en el nuevo formato de Proyectos** (bloque: el
  proyecto y su primer integrante comparten fila, los demás coautores van
  en filas separadas con los campos del proyecto en blanco): se agregó
  "arrastre" de los campos del proyecto hacia las filas de coautor en
  `cargar_proyectos`. Antes se perdían todos los coautores salvo el primero
  (175 → 230 registros correctos).
- **Rutas fijas desactualizadas** a un scrape viejo (`reports/excel`,
  diciembre 2025) en 4 archivos distintos
  (`vista_seguimiento_grupos.py`, `vista_clasificacion_minciencias.py`,
  `build_categorias_grupos_957.py`, `simulador_957.py`) — todas ahora usan
  `_carpeta_gruplac_mas_reciente()`.
- **Crash con archivos grandes concurrentes**: la extracción en paralelo
  (`ThreadPoolExecutor`) de los 6 extractores terminaba el proceso Python
  sin ninguna excepción capturable con los archivos actuales (CIARP +
  Informe Extensión, varios MB); secuencial no falla y tarda menos de 1s.
- Filtro de años en Seguimiento Grupos: el año "hasta" quedaba fijo en 2025,
  excluyendo silenciosamente todo 2026 de cada verificación.

### Cambiado
- **Panorama General retirado de la UI** (botón y punto de entrada
  comentados en `vista_seguimiento_grupos.py`) — el resumen "IA" era un
  solo prompt de una vía contra Ollama, no un chat interactivo; se
  reemplaza por el chatbot nuevo (ver arriba) cuando se decida reactivarlo.
- **Estadísticas 957** (primera versión, dentro de la app principal) se
  quitó y reconstruyó como `UI_clasificacion/` (separada, ver arriba) — el
  usuario prefirió una UI independiente para poder integrarla después sin
  arriesgar la app principal mientras se sigue ajustando.
- Inicio: se quitó la tabla de archivos y los contadores numéricos de
  "datos cargados" — ahora solo confirma que la carga fue exitosa.

### Agregado
- **Panorama General** (`DialogoPanoramaGeneral` en `vista_seguimiento_grupos.py`):
  tablero agregado sobre el caché de verificación — tarjetas resumen (grupos
  con datos, cumplimiento promedio, grupos en estado crítico, total
  faltantes), gráfico de % de cumplimiento por grupo (ordenado peor a
  mejor, coloreado por banda de estado) y gráfico de faltantes por
  categoría. Botón nuevo en el toolbar de Seguimiento Grupos.
- **Resumen narrativo con IA** dentro de Panorama General: botón "Generar
  resumen (IA)" que usa un modelo local de Ollama (`qwen2.5:3b`, streaming,
  hilo aparte) para redactar en prosa lo mismo que ya muestran las tarjetas
  y gráficos — no consulta datos nuevos ni sale a internet.

### Corregido
- Bug de rendimiento en el motor de comparación (`comparador_faltantes.py`):
  el filtro barato de tokens en común corría *después* del score costoso
  (`SequenceMatcher`) en vez de antes — "Verificar contra GrupLAC (nuevo)"
  pasó de una proyección de 7-8 horas a ~5 minutos, mismo resultado exacto.
- Extracción de título de productos GrupLAC sin numeración de ítem
  (Talleres de Creación, Eventos Artísticos, Procesos de apropiación
  social) — se quedaba con el bloque de texto crudo completo en vez del
  título, y productos que sí estaban en GrupLAC salían como "Faltante
  real" (caso real: "Laboratorio de creación en terracota bajo relieve",
  grupo LH).
- Truncado de nombre de hoja de Excel (límite de 31 caracteres) en el orden
  equivocado — "Procesos de apropiación social..." no calzaba en 87/127
  grupos pese a estar mapeada.
- Grupos con coma en su propio nombre (ej. "TERRITORIO, EDUCACIÓN Y
  SOCIEDAD") se partían en fragmentos inexistentes al separar celdas de
  Excel con varios grupos — sus productos quedaban invisibles para la
  comparación.
- Cobertura de mapeo hoja GrupLAC → categoría interna: se agregaron 16
  hojas de producto real que no estaban mapeadas ("Demás trabajos",
  "Obras o productos", "Signos distintivos", Programa Ondas, etc.) —
  cobertura ahora 73/73 hojas de producto encontradas en los 127 grupos.
- Fecha límite de carga hardcoded y vencida ("30 de noviembre de 2025") en
  el reporte Excel de faltantes — quitada, ahora solo muestra la fecha de
  generación real.

### Cambiado
- El popup "Duplicados" se simplificó: se quitó la pestaña de verificación
  de productos (ahora vive, ampliada, en el panel Cumplimiento — filtro de
  estado completo, columnas Responsable y Grupo encontrado, botones
  Recargar/Verificar/Exportar). Solo queda "Personas sin registro
  (GrupLAC)".
- Exportar a Excel desde Cumplimiento genera una sola hoja ("Faltantes
  Detalle") en vez de Resumen + una hoja por grupo.
- Se quitó el botón "Simular cat. 957" de Seguimiento Grupos (la clase
  `DialogoProductos957` queda en el código, sin usar, por si se retoma).

## 2026-07-09 (sesión anterior — ver `claude_1.md` para el detalle día a día)

- Falsos positivos de verificación de personas en GrupLAC corregidos
  (heurística de coincidencia de nombre demasiado laxa).
- Pestaña "Diff / Faltantes" embebida en Seguimiento Grupos (reemplazada
  después por el panel Cumplimiento ampliado, ver arriba).
- `comparador_gruplac_scrapeado.py`: primera versión del comparador contra
  el scraping nuevo de GrupLAC (antes se comparaba contra un snapshot
  estático, `gruplac_957.db`).
- Bug de datos real: columna "Estudiante" de Trabajos de Grado mostraba el
  número de documento en vez del nombre (columnas B/C invertidas al leer
  el Excel fuente).
