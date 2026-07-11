# Metodología del "Plan de Mejora 957" — qué es oficial y qué es aproximación propia

Este documento describe, con fórmulas, **de dónde sale cada número** que se
muestra en la pestaña *Plan de Mejora 957* (secciones 1 a 5), para poder
contrastarlo contra el documento oficial de la Convocatoria 957 de
MinCiencias y contra `data/output/medicion_957.xlsx` (el "paper"/medición
oficial que cubre 75 de los 125 grupos).

La idea central: **todo lo que se muestra como "valor oficial" viene
literalmente de `medicion_957.xlsx`, sin recalcular nada**. Lo único que
calculamos nosotros es la calibración del *simulador* (sección 5), que se
explica al final y está marcada explícitamente como aproximación.

---

## 1. El modelo de medición Conv. 957 (resumen)

MinCiencias agrupa los productos de un grupo de investigación en **5
categorías** y, a partir de ellas, calcula **7 indicadores agregados** que
determinan la categoría final del grupo en la escala:

```
D  →  C  →  B  →  A  →  A1
```

(de menor a mayor; `ORDEN_CATEGORIAS_MINCIENCIAS` en `src/constants.py`).

| Indicador (clave interna) | Significado MinCiencias | Sección en `medicion_957.xlsx` |
|---|---|---|
| `TOP`    | Nuevo Conocimiento TOP (productos de más alto impacto: A1) | `NC_TOP` |
| `TIPO_A` | Nuevo Conocimiento tipo A | `NC_A` |
| `TIPO_B` | Nuevo Conocimiento tipo B | `NC_B` |
| `AP`     | Apropiación Social del Conocimiento | `ASC` |
| `DPC`    | Divulgación Pública de la Ciencia | `DPC` |
| `FR_A`   | Formación de Recurso Humano A (doctorado/maestría) | *(sin desglose de productos)* |
| `FR_B`   | Formación de Recurso Humano B (especialización/pregrado) | *(sin desglose de productos)* |

Esta tabla es `CODIGO_957_A_INDICADOR` / `SECCION_957_POR_INDICADOR` en
`src/constants.py`. **FR_A y FR_B no tienen desglose por tipo de producto en
el documento oficial** — el valor del indicador existe, pero no hay una hoja
"productos" para esas dos secciones.

---

## 2. Fórmula λ por producto (oficial, verificada)

Cada producto contabilizado por MinCiencias dentro de una ventana de
vigencia aporta un **peso λ ("lambda")** según:

```
λ(subtipo) = ln( 1 + total / ventana )
```

donde:

- `total` = número de productos de ese subtipo que el grupo tiene dentro de
  la ventana de vigencia.
- `ventana` = número de años de vigencia de ese subtipo (5, 7 o 10 años
  según el tipo de producto).

**Esta fórmula fue validada empíricamente** comparando, para ~1.500 filas de
la hoja `productos` de `medicion_957.xlsx`, la columna `lambda_val` contra
`ln(1 + total/ventana)`: **coincide en el 95.3% de los casos** (la diferencia
en el resto se explica por redondeos / productos justo en el borde de la
ventana).

### 2.1 Ventanas de vigencia por subtipo (hoja `productos`)

La ventana es **la misma para todos los grupos** — depende solo del tipo de
producto. Algunos ejemplos por sección (lista completa: 106 subtipos en
`ventanas_subtipo`, cacheados en `_oficial_957`):

| Sección | Ventana = 5 años | Ventana = 7 años | Ventana = 10 años |
|---|---|---|---|
| `NC_TOP` (→ `TOP`) | Capítulo de libro A, A1 | Artículo A1/A2 (+Open Access), Libro A/A1 | Producción art./arq./diseño A, A1 |
| `NC_A` (→ `TIPO_A`) | Cap. libro B, innovación biológica, patente, software/prototipo, RNL_B, signos distintivos | Artículo B/C (+Open Access), Libro B | Prod. art./arq./diseño B, marcas/4 |
| `NC_B` (→ `TIPO_B`) | Cap. libro C, artículo D (+Open), libros formación, nuevos registros, ... | Libro C | Prod. art./arq./diseño C |
| `ASC` (→ `AP`) | Apropiación social (FIS, GPP_A/B/C, FCP_A/C, TCCG_A) | — | — |
| `DPC` (→ `DPC`) | 57 subtipos de divulgación (eventos, libros divulgativos, contenidos digitales, etc.) | — | — |

### 2.2 Ejemplo numérico real

Grupo "GESTIÓN DE SISTEMAS ELÉCTRICOS, ELECTRÓNICOS Y AUTOMÁTICOS" (sección
`NC_TOP`):

| Subtipo | Total actual | Ventana | λ = ln(1+total/ventana) |
|---|---|---|---|
| `CAP_LIB_A1` (Capítulo de libro A1) | 1 | 5 | ln(1+1/5) = **0.1823** |
| `ART_A1` (Artículo Q1) | 1 | 7 | ln(1+1/7) = **0.1335** |
| `ART_OPEN_A2` | 2 | 7 | ln(1+2/7) = **0.2513** |
| `ART_OPEN_A1` | 3 | 7 | ln(1+3/7) = **0.3567** |
| ... (resto de subtipos `NC_TOP` en 0) | 0 | — | 0 |
| **Σ λ (`lambda_actual_total`)** | | | **0.9238** |

---

## 3. De λ al "valor del indicador" — qué sabemos y qué no

La hoja `indicadores` de `medicion_957.xlsx` trae, para cada grupo y cada
sección (`NC_TOP`, `NC_A`, `NC_B`, `ASC`, `DPC`, `FRH_A`, `FRH_B`), un
**`valor_indicador` ya calculado por MinCiencias**.

**Lo que NO hemos podido reproducir**: la fórmula exacta que convierte la
suma de λ de los productos de un grupo (Σλ ≈ 0.92 en el ejemplo anterior) en
ese `valor_indicador` (743.711 para `TOP` en el mismo grupo). La proporción
entre ambos números (`743.711 / 0.9238 ≈ 805`) es **enorme** y, lo más
importante, **distinta para cada grupo e indicador** (en el mismo grupo:
≈805 para `TOP`, ≈258 para `TIPO_A`, ≈194 para `TIPO_B`, ≈2000 para `AP`).
Esto sugiere que la fórmula real de MinCiencias incluye al menos uno de:

- una normalización/ponderación por el tamaño total del SNCTeI o del área de
  conocimiento (no disponible en `medicion_957.xlsx`, que solo trae
  resultados ya agregados),
- pesos por subtipo de producto distintos de 1 (es decir, no es una simple
  Σλ sino Σ(wᵢ·λᵢ) con pesos wᵢ que no están documentados en la hoja),
- y/o una escala/transformación adicional definida en el documento oficial
  de la convocatoria (anexo técnico) que no viene en el Excel de resultados.

**Conclusión práctica**: para las secciones 1–4 del Plan de Mejora **no
recalculamos nada** — usamos `valor_indicador` y los cuartiles tal cual
vienen en `medicion_957.xlsx`, que son 100% oficiales. La incógnita de la
fórmula de agregación solo afecta al **simulador (sección 5)**, ver §5.

---

## 4. Cuartiles y "mínimo nacional" (Sección 4 — 100% oficial)

La hoja `cuartiles` de `medicion_957.xlsx` trae, para cada indicador y cada
combinación (categoría, área de conocimiento), la **distribución nacional**
de `valor_indicador` entre los grupos que están en esa categoría/área:

```
min ≤ q4 ≤ q3 ≤ q2 ≤ max
```

(`min`=percentil 0, `q4`≈percentil 25, `q3`≈percentil 50, `q2`≈percentil 75,
`max`=percentil 100 — son los **cuartiles de los grupos pares**, no del
grupo analizado).

### 4.1 ¿Qué columna es el "requisito mínimo"?

Para que un grupo ascienda de categoría actual → categoría objetivo, su
`valor_indicador` debe alcanzar **al menos** la columna de la distribución de
los grupos que YA están en la categoría objetivo, según
`CUARTIL_OBJETIVO_POR_CATEGORIA`:

| Categoría objetivo | Columna exigida | Interpretación |
|---|---|---|
| `C`  | `min` | Superar el mínimo histórico de los grupos C |
| `B`  | `q4`  | Estar sobre el percentil 25 de los grupos B |
| `A`  | `q3`  | Estar sobre la mediana de los grupos A |
| `A1` | `q2`  | Estar sobre el percentil 75 de los grupos A1 |

Este mapeo fue **validado empíricamente** comparando, para distintos grupos y
categorías, qué columna hace que `cumple = (tu_valor ≥ requisito)` sea
consistente con la categoría real asignada por MinCiencias.

### 4.2 Ejemplo numérico real

Grupo "GESTIÓN..." (categoría actual `A`, objetivo `A1`, área "Ingeniería y
Tecnología", columna objetivo = `q2`):

| Indicador | Tu valor | min | q4 | q3 | **q2 (requisito)** | max | ¿Cumple? |
|---|---|---|---|---|---|---|---|
| TOP    | 743.71  | 73.4  | 471.5 | 1138.7 | **2155.98** | 7098.2 | ❌ |
| TIPO_A | 246.16  | 20.0  | 257.5 | 536.0  | **1049.07** | 7103.9 | ❌ |
| TIPO_B | 340.09  | 5.5   | 51.0  | 160.0  | **420.02**  | 3539.9 | ❌ |
| AP     | 1386.29 | 109.4 | 364.6 | 940.0  | **1941.56** | 8968.3 | ❌ |
| DPC    | 3284.85 | 109.4 | 364.6 | 940.0  | **1941.56** | 8968.3 | ✅ |
| FR_A   | 477.76  | 91.2  | 235.0 | 376.0  | **684.21**  | 2480.0 | ❌ |
| FR_B   | 3679.69 | 18.2  | 435.3 | 769.9  | **1272.56** | 4065.3 | ✅ |

Todos estos números salen **tal cual** de `medicion_957.xlsx` (hojas
`indicadores` y `cuartiles`) vía `_indicadores_oficiales` /
`_cuartiles_oficiales`. La sección 4 de la UI es una transcripción directa de
esta tabla.

---

## 5. Simulador (Sección 5) — metodología propia (NO oficial, sí calibrada con datos oficiales)

Como se explicó en §3, no existe (a nuestro alcance) una fórmula pública que
diga "si agregas 1 artículo Q1 más, tu `valor_indicador` sube en X". El
simulador resuelve esto con una **calibración lineal local, propia de cada
grupo y cada indicador**, ancorada a los valores reales de ese grupo en
`medicion_957.xlsx`.

### 5.1 Factor de conversión (`ratio`)

Para cada indicador con brecha (no cumple el requisito de §4):

```
ratio = valor_indicador_actual / Σλ_actual
```

donde `Σλ_actual` es la suma de `λ(subtipo) = ln(1+total/ventana)` (§2) sobre
**todos los subtipos de la sección correspondiente que el grupo ya tiene
registrados** en `medicion_957.xlsx`.

> Importante: `ratio` es **diferente para cada grupo e indicador** — es
> precisamente el factor "desconocido" de §3, pero en vez de adivinarlo lo
> medimos directamente con los dos únicos puntos de datos que sí tenemos
> (Σλ y `valor_indicador`, ambos del mismo grupo). Por construcción, con
> `n=0` adiciones el simulador reproduce **exactamente** `tu_valor` de la
> sección 4 (no hay desajuste de escala).

### 5.2 Proyección al simular nuevas unidades

Si el usuario indica que va a generar `n` unidades adicionales de un subtipo
que hoy tiene `total_actual`:

```
λ_nuevo(subtipo)  = ln(1 + (total_actual + n) / ventana)
Δλ                = λ_nuevo(subtipo) − λ_actual(subtipo)
Σλ_simulado       = Σλ_actual + Σ Δλ   (sumando sobre todos los subtipos con n>0)

valor_simulado    = ratio × Σλ_simulado
```

`valor_simulado` queda en la **misma escala** que `tu_valor` / `requisito
mínimo` de la sección 4, y `¿Cumples?` se evalúa igual que en §4
(`valor_simulado ≥ requisito_minimo`).

### 5.3 Ejemplo numérico real

Para el indicador `TOP` del grupo "GESTIÓN...": `ratio ≈ 805.02`,
`Σλ_actual ≈ 0.9238` (de la tabla de §2.2), `requisito_minimo ≈ 2155.98`.

Si el usuario simula generar **+5 capítulos de libro A1** (`CAP_LIB_A1`,
`ventana=5`, `total_actual=1`):

```
λ_nuevo  = ln(1 + (1+5)/5) = ln(2.2) = 0.7885
Δλ       = 0.7885 − 0.1823 = 0.6061
Σλ_sim   = 0.9238 + 0.6061 = 1.5299

valor_simulado = 805.02 × 1.5299 ≈ 1231.7   →  todavía < 2155.98  → ❌
```

(coincide con el resultado calculado por `_recalcular_simulador` en la UI).

### 5.4 Supuestos y límites de esta calibración

- **Supone proporcionalidad lineal** entre Σλ y `valor_indicador` alrededor
  del punto actual del grupo (`ratio` constante). Es razonable para
  variaciones moderadas, pero **no está garantizado para extrapolaciones muy
  grandes** (si la fórmula real de MinCiencias no es lineal en Σλ, el `ratio`
  real podría cambiar a medida que el grupo crece mucho).
- Si `Σλ_actual = 0` (el grupo no tiene ningún producto registrado en esa
  sección en `medicion_957.xlsx`), `ratio` es indefinido y el simulador
  muestra el mensaje "no se puede calibrar" en vez de un número.
- `FR_A`/`FR_B` no tienen desglose de productos oficiales → se muestra
  "MinCiencias no desglosa este indicador..." y no hay simulador para esos
  dos indicadores.
- Los candidatos a simular se limitan a 6 por indicador, priorizando los
  subtipos que el grupo **ya produce** (más realista que sugerir un tipo de
  producto completamente nuevo).

---

## 6. Grupos sin datos oficiales (fuente `bd_interna`)

50 de los 125 grupos no aparecen en `medicion_957.xlsx`. Para esos grupos:

- La **categoría actual y área** vienen de
  `data/cache/categorias_grupos_957.json` (generado por
  `build_categorias_grupos_957.py` a partir de la hoja "Datos Básicos" de
  cada Excel GrupLAC: campo "Clasificación").
- Los **7 indicadores** se recalculan desde la BD interna
  (`academia_utp_integrado.db`) usando `VENTANAS_957` (`src/constants.py`):
  cada producto suma un **λ fijo** (no `ln(1+total/ventana)`, sino una
  constante por categoría de publicación: A1→1.00, A2→0.75, B→0.50, C/D→
  0.25/0.10, libro→2.00, capítulo→1.00, etc.) si está dentro de su ventana.
- **No hay sección 4 ni 5** (cuartiles/simulador) para estos grupos, porque
  no existe una distribución nacional con la que compararlos en
  `medicion_957.xlsx`. La UI lo indica explícitamente
  ("Fuente de indicadores: BD interna...").
- Esta vía (`VENTANAS_957`) es una **aproximación nuestra** anterior a la
  calibración de §5, y **no** está validada contra `medicion_957.xlsx` (no
  hay datos oficiales contra qué comparar esos 50 grupos).

---

## 7. Resumen — oficial vs. aproximado

| Elemento | Sección UI | Origen | Estatus |
|---|---|---|---|
| Categoría actual, área | Estado | `medicion_957.xlsx` (75) / `categorias_grupos_957.json` (50) | Oficial / oficial (otro documento GrupLAC) |
| `valor_indicador` actual | 1, 4 | hoja `indicadores` | **100% oficial** |
| Umbral percentil-25 grupos del área en cat. objetivo | 1 | calculado sobre `indicadores` oficiales | Oficial (cálculo estadístico simple sobre datos oficiales) |
| min/q4/q3/q2/max por indicador | 4 | hoja `cuartiles` | **100% oficial** |
| Requisito mínimo (columna según categoría objetivo) | 4 | `CUARTIL_OBJETIVO_POR_CATEGORIA` | Mapeo validado empíricamente, sobre datos oficiales |
| λ = ln(1+total/ventana) por subtipo | 5 (interno) | fórmula validada al 95.3% vs columna `lambda_val` de `productos` | Oficial (verificada) |
| `ratio` = valor_indicador / Σλ | 5 | **calibración propia**, por grupo/indicador | Aproximación (lineal local), anclada a datos oficiales |
| `valor_simulado` = ratio × Σλ_simulado | 5 | **proyección propia** | Aproximación, ver límites en §5.4 |
| Indicadores de grupos `bd_interna` (50/125) | 1 | `VENTANAS_957` (λ fijo por categoría) | Aproximación propia, sin contraste oficial |

---

## 8. Archivos fuente relevantes

- `data/output/medicion_957.xlsx` — hojas `grupos`, `indicadores`,
  `productos`, `cuartiles` (documento de medición oficial, 75/125 grupos).
- `data/cache/categorias_grupos_957.json` — respaldo para los 50 grupos
  restantes (`src/build_categorias_grupos_957.py`).
- `src/constants.py` — `INDICADORES_957`, `ORDEN_CATEGORIAS_MINCIENCIAS`,
  `CODIGO_957_A_INDICADOR`, `SECCION_957_POR_INDICADOR`,
  `INDICADOR_957_A_CUARTIL`, `CUARTIL_OBJETIVO_POR_CATEGORIA`, `VENTANAS_957`.
- `src/analisis_seguimiento.py` (`SimuladorCategoriaInterna`):
  - `_cargar_datos_oficiales_957`, `_indicadores_oficiales`,
    `_productos_oficiales`, `_cuartiles_oficiales` → datos oficiales (§3-4).
  - `opciones_simulador`, `_etiqueta_subtipo` → calibración del simulador
    (§5).
  - `analizar_brechas_area` → orquesta todo y devuelve el diccionario `plan`
    que consume `src/views/vista_seguimiento_grupos.py`.
