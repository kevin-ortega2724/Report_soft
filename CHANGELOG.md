# Changelog

Formato libre, en español, más reciente primero. Para el detalle técnico de
los bugs del pipeline de verificación ver
[`docs/VERIFICACION_GRUPLAC.md`](docs/VERIFICACION_GRUPLAC.md) — acá solo el
resumen de qué cambió y por qué.

## 2026-07-11

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
