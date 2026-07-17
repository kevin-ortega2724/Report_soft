# UI_clasificacion — Clasificación 957 (vigente + proyección estimada)

UI **separada** de la app principal (`run.py` / `main_10.py`) a propósito —
se integrará más adelante, pero por ahora se prueba y evoluciona de forma
independiente. Todo el código de esta carpeta importa lo ya construido y
validado en `src/` (no reimplementa nada).

## Qué muestra

Por cada uno de los 127 grupos scrapeados de GrupLAC:

1. **Categoría vigente**: la que reporta MinCiencias en el perfil público
   del grupo (hoja "Datos básicos" → "Clasificación"). 100% técnico
   (scraping), siempre tan actualizado como la última corrida de
   "Actualizar GrupLAC (Web)" en la app principal. No requiere ningún
   documento oficial adicional.
2. **Categoría oficial (base)** y **categoría proyectada (estimada)** —
   solo para los ~67 grupos que sí aparecen en `data/output/medicion_957.xlsx`
   (el documento de medición oficial de la Convocatoria 957). La proyección
   estima cómo quedaría el grupo si la convocatoria se recalculara HOY,
   usando su producción de artículos publicada DESPUÉS del cierre de la
   ventana oficial (31 dic 2023).

Al seleccionar un grupo en la tabla, el panel "Detalle del grupo" muestra el
cálculo metodológico completo: Índice Global y cuartil (Pasos 5-6),
indicadores contra los umbrales min/q4/q3/q2/max reales del área (100%
oficiales), las condiciones de cada categoría (A1/A/B/C, transcritas
literalmente de los PDF de MinCiencias) marcadas ✓/✗, y el detalle
artículo-por-artículo usado en la proyección (ISSN, año, categoría
Publindex encontrada, λ antes/después).

## Cómo correrla

```powershell
.venv\Scripts\python.exe UI_clasificacion\run_clasificacion.py
```

Requiere que ya existan `data/db/academia_utp_integrado.db` y al menos una
carpeta `data/reporte excel_<fecha>` (generadas por la app principal).

## Archivos

| Archivo | Qué hace |
|---|---|
| `obtener_publindex.py` | Descarga (una vez) el dataset abierto oficial "Revistas Indexadas, Índice Nacional Publindex" (datos.gov.co, Socrata, id `mwmn-inyg`) y lo cachea en `data_sim/publindex.csv`. Correr de nuevo si se necesita refrescar. |
| `proyeccion_957.py` | Cruza artículos internos nuevos (ISSN + año > 2023) contra Publindex para saber su calidad real (A1/A2/B/C), arma los `ajustes` por subtipo y llama a `Simulador957.proyectar_productos()` (ya validado en `src/simulador_957.py`). |
| `vista_clasificacion.py` | `VistaClasificacion957(QWidget)` — la UI: tabla, chips resumen, gráfica de distribución, y el panel de detalle metodológico por grupo. |
| `run_clasificacion.py` | Entry point standalone (`QApplication` + `QMainWindow` propios). |
| `data_sim/publindex.csv` | Cache local del dataset Publindex (~6276 filas, se versiona en git — es dato público, no personal). |
| `data_sim/Ficha-de-Asesoria-*.xlsx` | Documento real de MinCiencias que sirvió para validar la metodología (contiene datos personales de un docente — **no se versiona**, excluido por `*.xlsx` en `.gitignore`). |

## Por qué existe (contexto para retomar)

1. Ya existía, huérfana y sin punto de entrada en la app principal, una
   clase `DialogoProductos957` (en `src/views/vista_seguimiento_grupos.py`,
   ~línea 881) con un "Plan de Mejora 957" completo (Diagnóstico,
   Recomendaciones, Mínimos nacionales, Simulador) construido sobre
   `Simulador957`. Se decidió NO reactivarla y en su lugar construir esta UI
   nueva y separada, reusando el mismo motor validado.
2. `Simulador957.simular()` (modo verificación) está calibrado al 100%
   contra los 60 grupos oficiales con categoría conocida — es fiable.
   `Simulador957.proyectar_productos()` también está validado (con
   ajustes=0 reproduce exacto el valor oficial).
3. El obstáculo real para "proyectar con datos actuales": el scraping
   público de GrupLAC NO trae la calidad Publindex (A1/A2/B/C) de un
   artículo — solo título/revista/ISSN/año. Se investigó y confirmó que
   Publindex SÍ es público como dataset abierto (ver `obtener_publindex.py`),
   lo que permitió construir la proyección real en vez de descartarla.

## Limitaciones conocidas (documentadas a propósito, no ocultas)

- **Solo se proyectan artículos** (NC_TOP/NC_A/NC_B vía ART_A1/A2/B/C/D).
  Libros, capítulos, extensión (ASC), divulgación (DPC) y formación de
  recurso humano (FRH_A/FRH_B) quedan en su valor oficial, sin proyectar —
  el campo interno `publicaciones.categoria` no sirve para clasificar
  libros/capítulos (mezcla tipos descriptivos con códigos de incentivos
  internos UTP/CIARP, ver `proyeccion_957.py`).
- **Publindex (dataset abierto) no tiene vigencia 2023/2024 todavía** — tope
  real 2022. Artículos posteriores usan la vigencia más reciente disponible
  para esa revista como aproximación documentada (`TOPE_VIGENCIA_PUBLINDEX`
  en `proyeccion_957.py`).
- **Los ~60 grupos sin línea base oficial** (no están en `medicion_957.xlsx`)
  solo muestran categoría vigente, sin proyección — no hay `ratio` que
  calibrar sin un valor oficial de referencia.
- La condición "investigador emérito/sénior/asociado vinculado" (una de las
  condiciones oficiales de categoría) nunca se evalúa — ese dato requiere el
  CvLAC individual de cada investigador, no existe en el scraping de GrupLAC
  a nivel de grupo. Documentado en `simulador_957.py`.

## Próximos pasos posibles (no hechos, ideas para retomar)

- Integrar esta UI como pestaña de la app principal (`main_10.py`) cuando se
  decida — el patrón de integración ya se usó y revirtió una vez con
  `chatbot_investigacion.py` (ver `main_10.py`, comentarios en
  `_tab_registry`), así que el mismo patrón mínimo (import + entrada en
  `_tab_titles`/`_tab_registry`) aplica acá.
- Si aparece una fuente confiable de clasificación de libros/capítulos, o
  Publindex publica vigencias 2023+, ampliar `proyeccion_957.py` sin tocar
  la UI (ya está separado en su propio módulo).
- `obtener_publindex.py` se puede volver a correr cuando se quiera refrescar
  el cache (no hay lógica de expiración automática todavía).
