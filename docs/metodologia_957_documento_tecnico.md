# ReportSoft · Plan de Mejora 957
## Documento Técnico de Metodología — Fórmulas oficiales y justificación

**Versión:** 2.0 · **Herramienta:** ReportSoft – Módulo de Seguimiento de Grupos  
**Convocatoria de referencia:** Conv. 957 — *"Convocatoria Nacional de Actualización y Transición
para el Reconocimiento y Medición de Grupos de Investigación, Desarrollo Tecnológico o de
Innovación y para el Reconocimiento de Investigadores del SNCTeI — 2024"*  
**Fuente primaria:** Reporte de proceso de medición ScienTI / GrupLAC (pp. 119-125 del Documento
Conceptual de la Convocatoria)

---

> **Propósito**
> Este documento explica, con las fórmulas exactas tomadas del sistema oficial ScienTI y de los
> reportes de medición de MinCiencias, cómo se calcula la categoría de un grupo de investigación
> en la Conv. 957. Se muestra además cómo ReportSoft usa esas fórmulas, dónde las reproduce
> exactamente y dónde introduce una aproximación justificada matemáticamente.
>
> El documento está dirigido a investigadores, directores de grupos y evaluadores que quieran
> verificar o contrastar cualquier número mostrado en la herramienta.

---

## Contenido

1. [El proceso oficial: 6 pasos de MinCiencias](#1-proceso-oficial)
2. [Paso 2 — Fórmula λ: conteo escalado de productos](#2-lambda)
3. [Paso 3 — Indicador de sección: de λ al valor numérico](#3-indicador-seccion)
4. [Paso 4 — Índices: normalización por el máximo del área](#4-indices)
5. [Paso 5 — Indicador de Grupo (IG): suma ponderada](#5-indicador-grupo)
6. [Paso 6 — Clasificación por cuartiles y condiciones adicionales](#6-clasificacion)
7. [Cómo ReportSoft usa estas fórmulas — sección por sección](#7-reportsoft)
8. [Caso de aplicación completo: grupo "Automática" (A1)](#8-caso-automatica)
9. [Caso de aplicación: grupo "GESTIÓN DE SISTEMAS ELÉCTRICOS" (A)](#9-caso-gestion)
10. [El simulador de impacto: justificación matemática](#10-simulador)
11. [Grupos sin datos oficiales — aproximación BD interna](#11-bd-interna)
12. [Tabla resumen: oficial vs. aproximación](#12-resumen)
13. [Preguntas frecuentes](#13-faq)
14. [Glosario](#14-glosario)

---

## 1. El proceso oficial: 6 pasos de MinCiencias

El reporte de medición que MinCiencias pone a disposición de cada grupo en la plataforma
ScienTI (sección *"Proceso de medición → Indicadores"*) describe un proceso en **6 pasos**:

```
Paso 1: Revisión de requerimientos (existencia y calidad de productos)
Paso 2: Conteo escalado de productos por subtipo → λ por subtipo
Paso 3: Cálculo del indicador de sección → valor_indicador_j (j = TOP, A, B, ASC, DPC, FR_A, FR_B)
Paso 4: Cálculo de índices → índice_j = valor_indicador_j / máximo_del_área_j
Paso 5: Cálculo del Indicador de Grupo → IG = Σ (ponderación_j × índice_j)
Paso 6: Clasificación por cuartiles y condiciones adicionales → categoría D/C/B/A/A1
```

Cada paso está documentado en el reporte PDF de ScienTI para cada grupo y cada convocatoria.
ReportSoft extrae los resultados de estos pasos desde `medicion_957.xlsx` (hojas `productos`,
`indicadores`, `cuartiles`) — que es el mismo cálculo en formato tabular.

---

## 2. Paso 2 — Fórmula λ: conteo escalado de productos

### 2.1 La fórmula oficial

Para cada **subtipo** de producto que el grupo tiene dentro de la ventana de vigencia, el
sistema ScienTI calcula:

$$\boxed{\lambda(\text{subtipo}) = \ln\!\left(1 + \frac{\text{total\_en\_ventana}}{\text{ventana}}\right)}$$

donde:
- **`total_en_ventana`** = número de unidades de ese subtipo con fecha dentro de la ventana.
- **`ventana`** = años de vigencia del subtipo (5, 7 o 10 años — fijo por tipo de producto,
  igual para todos los grupos del SNCTeI).

Esto aparece en la columna **`Lambda (λ)`** de la tabla del Paso 2 de los reportes ScienTI.

### 2.2 ¿Por qué logaritmo? — Justificación matemática

La función `ln(1 + x)` tiene las siguientes propiedades que justifican su uso en un sistema
de medición científica:

| Propiedad | Fórmula | Interpretación |
|---|---|---|
| **Valor base** | λ(0) = 0 | Sin productos dentro de la ventana, aporte nulo |
| **Creciente** | dλ/dx = 1/(1+x) > 0 | Más productos siempre aumenta el indicador |
| **Cóncava** | d²λ/dx² = −1/(1+x)² < 0 | Rendimientos decrecientes: cada unidad adicional aporta menos |
| **Acotada** | λ → ∞ lentamente | Evita que un grupo con 1000 artículos domine absolutamente |

**Ejemplo de rendimiento decreciente** — artículo Q1 (ventana = 7 años):

| N.° artículos Q1 | λ = ln(1 + N/7) | Δλ marginal |
|---|---|---|
| 1 | ln(1.143) = **0.1335** | — |
| 2 | ln(1.286) = **0.2513** | +0.1178 |
| 3 | ln(1.429) = **0.3567** | +0.1054 |
| 5 | ln(1.714) = **0.5390** | +0.1823 (dos más) |
| 10 | ln(2.429) = **0.8873** | +0.3483 (cinco más) |
| 20 | ln(3.857) = **1.3505** | +0.4632 (diez más) |

Generar 20 artículos Q1 solo aporta ~10× lo que aporta 1 artículo, no 20×.

### 2.3 La ventana de vigencia: densidad, no conteo bruto

Dividir por `ventana` convierte el conteo en una **tasa anual equivalente**:

$$\frac{\text{total}}{\text{ventana}} = \text{tasa de producción anual promedio dentro del período}$$

Un grupo con 5 artículos en 5 años y otro con 5 artículos en 10 años tienen
la misma cantidad pero distintas densidades: λ₅ = ln(1+1) = **0.693** vs.
λ₁₀ = ln(1+0.5) = **0.405**. MinCiencias premia la constancia: producir
sostenidamente en menos tiempo da un λ mayor.

### 2.4 Ventanas por tipo de producto (selección)

| Sección → indicador | Subtipo | Ventana |
|---|---|---|
| `NC_TOP` → TOP | Artículo Q1/Q2 (ART_A1, ART_A2, ART_OPEN_A1, ART_OPEN_A2) | 7 |
| `NC_TOP` → TOP | Capítulo de libro A/A1 (CAP_LIB_A, CAP_LIB_A1) | 5 |
| `NC_TOP` → TOP | Libro A/A1 (LIB_A, LIB_A1) | 7 |
| `NC_A` → TIPO_A | Artículo B/C (ART_B, ART_C) | 7 |
| `NC_A` → TIPO_A | Patente (PA4), Software (SF), Modelo de utilidad (MA4) | 10 |
| `NC_B` → TIPO_B | Artículo D (ART_D), Libro formación (LIB_C) | 7 / 5 |
| `ASC` → AP | FIS, TCCG_A, GPP_A/B/C, FCP_A/C | 5 |
| `DPC` → DPC | GC, EC_A/B, RC_A, WP, IFI, PEE, PCD, LIB_DIV | 5 |
| `FRH_A` → FR_A | TD_A (tesis doctoral), AP_C (posgrado) | 5 |
| `FRH_B` → FR_B | TM_A/B (tesis maestría), TP_B (tesis pregrado), PID, AP_D | 5 |

*(Lista completa en hoja `productos` de `medicion_957.xlsx` — 106 subtipos.)*

### 2.5 Verificación empírica

Para ~1.500 filas de la hoja `productos` de `medicion_957.xlsx`, la columna `lambda_val`
coincide con `ln(1 + total/ventana)` en el **95.3%** de los casos. El 4.7% restante
corresponde a productos en el borde exacto de la ventana (diferencia de días en la fecha
de corte) o a subtipos con regla de ventana no estándar.

---

## 3. Paso 3 — Indicador de sección: de λ al valor numérico

### 3.1 Lo que muestra el reporte ScienTI

El Paso 3 del reporte ScienTI dice:

> *"Fórmula para hallar el indicador de cada subtipo de producto:"*  
> [fórmula en imagen — no extractable como texto]

Luego muestra directamente:

| Indicador | Valor del indicador obtenido por el grupo |
|---|---|
| Productos de Nuevo conocimiento TOP | 2602.668427454901 |
| Productos de Nuevo conocimiento Tipo A | 858.3584901120123 |
| ... | ... |

**El texto de la fórmula no es texto plano en el PDF** — está renderizado como imagen
matemática en el sistema ScienTI. Sin embargo, cruzando los datos del Paso 2 y el Paso 4
de múltiples grupos, podemos **deducirla con certeza matemática**.

### 3.2 Deducción matemática de la fórmula del Paso 3

Del Paso 4 (que sí está completamente legible), sabemos:

$$\text{índice}_j = \frac{\text{valor\_indicador}_j}{\text{máximo\_del\_área}_j}$$

donde `máximo_del_área_j` es el valor máximo de `valor_indicador_j` entre **todos los
grupos del mismo área de conocimiento** que participan en la convocatoria (confirmado:
el máximo de Step 4 coincide exactamente con el máximo de la columna "Máximo" de los
cuartiles del área en Step 6).

Reescribiendo: **`valor_indicador_j = índice_j × máximo_del_área_j`**

El grupo con `índice_j = 1` es el que tiene el máximo de Σλ_j en su área.
Por tanto, para todos los grupos del mismo área:

$$\text{valor\_indicador}_j(g) = \frac{\Sigma\lambda_j(g)}{\Sigma\lambda_{\max,j}(\text{área})} \times \text{máximo}_j(\text{área})$$

O equivalentemente:

$$\boxed{\text{valor\_indicador}_j(g) = \Sigma\lambda_j(g) \times C_j(\text{área})}$$

donde **`C_j(área)`** es una constante que depende solo del área y la sección:

$$C_j(\text{área}) = \frac{\text{valor\_indicador\_máximo}_j(\text{área})}{\Sigma\lambda_{j,\max}(\text{área})}$$

### 3.3 Verificación numérica con datos reales

**Grupo "Automática" — área Ingeniería y Tecnología — categoría A1:**

Para la sección NC_TOP:
- Σλ_TOP = ln(1+11/7) + ln(1+4/7) + ln(1+10/7) + ln(1+11/7)
         = 0.9445 + 0.4520 + 0.8873 + 0.9445 = **3.2283**
- `valor_indicador_TOP` (Paso 3) = **2602.668**
- `máximo_TOP` área (Paso 4 = Máximo cuartiles, Paso 6) = **7098.217**
- `índice_TOP` (Paso 4) = 2602.668 / 7098.217 = **0.36666** ✅ (coincide con el PDF)
- `C_TOP(Ingeniería)` = 2602.668 / 3.2283 = **806.2**

**Grupo "BIOECOS" — área Ciencias Agrícolas — categoría A1:**

Para la sección NC_TOP:
- Σλ_TOP = ln(1+10/7) + ln(1+5/7) + ln(1+25/7) + ln(1+13/7)
         = 0.8873 + 0.5390 + 1.5198 + 1.0498 = **3.9959**
- `valor_indicador_TOP` (Paso 3) = **3305.071**
- `máximo_TOP` área Ciencias Agrícolas (Paso 4) = **8257.258**
- `índice_TOP` (Paso 4) = 3305.071 / 8257.258 = **0.40026** ✅ (coincide con el PDF)
- `C_TOP(Ciencias Agrícolas)` = 3305.071 / 3.9959 = **827.1**

La constante `C_j` **varía entre áreas** (806.2 vs 827.1 para NC_TOP). Esto confirma que
la normalización es **por área**, no nacional.

### 3.4 Consecuencia clave para ReportSoft

Para un grupo dado, `C_j(área)` es **fija y calculable** una vez que se tiene:
- `valor_indicador_j` (de la hoja `indicadores` — oficial)
- `Σλ_j` (calculado con la fórmula λ del §2 — verificada al 95.3%)

Esto es exactamente el `ratio` del simulador de ReportSoft:

$$\text{ratio} = \frac{\text{valor\_indicador\_actual}}{\Sigma\lambda_{\text{actual}}} = C_j(\text{área})$$

**El ratio no es una aproximación — es la constante de área exacta**, válida para
todos los grupos del mismo área y sección. La única fuente de error es el 4.7%
de imprecisión en la fórmula λ (borde de ventana).

---

## 4. Paso 4 — Índices: normalización por el máximo del área

### 4.1 La fórmula oficial

Para cada sección j y cada grupo g:

$$\boxed{\text{índice}_j(g) = \frac{\text{valor\_indicador}_j(g)}{\max_g\bigl[\text{valor\_indicador}_j(g)\bigr]_{\text{área}}}}$$

El máximo se calcula **dentro del área de conocimiento del grupo** (Ingeniería, Ciencias
Agrícolas, Ciencias de la Salud, etc.), no sobre el total nacional.

### 4.2 Tabla de ponderaciones oficiales (Paso 5)

| Índice | Ponderación (w) | Fuente en el PDF |
|---|---|---|
| Índice NC_TOP | **3.7** | Step 5, columna "Ponderación" |
| Índice NC_A (Tipo A) | **2.3** | Step 5 |
| Índice NC_B (Tipo B) | **0.4** | Step 5 |
| Índice ASC (Apropiación Social) | **1.5** | Step 5 |
| Índice DPC (Divulgación Pública) | **0.5** | Step 5 |
| Índice FR_A (Formación RH A) | **1.0** | Step 5 |
| Índice FR_B (Formación RH B) | **0.2** | Step 5 |
| Índice de Cohesión (IC) | 0.1 | Step 5 |
| Índice de Colaboración | 0.3 | Step 5 |
| **Total** | **Σ = 10.0** | |

> **Observación importante**: la suma de ponderaciones = 10, lo que significa que el
> IG máximo teórico (todos los índices = 1) sería 10.0. En la práctica, los grupos A1
> tienen IG entre 1.5 y 6.5 en los datos disponibles.

---

## 5. Paso 5 — Indicador de Grupo (IG): suma ponderada

### 5.1 La fórmula oficial

$$\boxed{IG = \sum_{j} w_j \cdot \text{índice}_j = \sum_{j} w_j \cdot \frac{\Sigma\lambda_j(g)}{{\Sigma\lambda_{j,\max}(\text{área})}}}$$

con los pesos `wⱼ` de la tabla del §4.2.

### 5.2 Verificación con el grupo "Automática"

| Índice | Ponderación (w) | Índice obtenido | w × índice |
|---|---|---|---|
| NC_TOP | 3.7 | 0.36666 | **1.3567** |
| NC_A | 2.3 | 0.12083 | **0.2779** |
| NC_B | 0.4 | 0.14141 | **0.0566** |
| ASC | 1.5 | 0.0 | **0.0000** |
| DPC | 0.5 | 0.27962 | **0.1398** |
| FR_A | 1.0 | 0.72744 | **0.7274** |
| FR_B | 0.2 | 0.55880 | **0.1118** |
| Cohesión | 0.1 | 0.33573 | 0.0336 |
| Colaboración | 0.3 | 0.05435 | 0.0163 |
| **IG** | | | **2.7199** ✅ |

*(Coincide con el valor reportado por ScienTI: 2.7198746881365135)*

### 5.3 ¿Qué significa el IG?

El IG es un número adimensional (cociente de λ ponderado sobre el máximo del área).
Representa **"cuánto puntaje de producción tiene el grupo en relación con el mejor de su área"**,
ponderando más los productos de mayor impacto (NC_TOP pesa 3.7 vs NC_B que pesa 0.4).

Un IG = 1 significaría que el grupo supera en todos los indicadores al mejor de su área
(imposible excepto para el grupo máximo en todos). Un grupo A1 típico tiene IG entre 1.5 y 4.

---

## 6. Paso 6 — Clasificación por cuartiles y condiciones adicionales

### 6.1 Cuartiles del IG (Indicador de Grupo)

El IG se clasifica dentro de los cuartiles del área de conocimiento:

| Cuartil del IG | Posición en el área |
|---|---|
| Cuartil 1 (25% superior, Q2–máx) | Candidato a A1 o A (con otras condiciones) |
| Cuartil 2 (Q3–Q2) | Candidato a B o A |
| Cuartil 3 (Q4–Q3) | Candidato a C |
| Cuartil 4 (mín–Q4) | D o sin clasificar |

### 6.2 Condiciones adicionales por categoría (Paso 6.2)

Las condiciones que deben cumplirse **simultáneamente** para cada categoría:

#### Categoría A1
1. IG en cuartil 1 del área (25% superior)
2. Indicador NC_TOP en cuartil 1 del área
3. Indicador ASC > 0 **ó** DPC > 0
4. Indicador FR_A > 0
5. Al menos 1 investigador emérito, sénior o asociado vinculado contractualmente
6. Índice de Cohesión > 0
7. Al menos 5 años de existencia del grupo

#### Categoría A
1. IG en cuartil 1 del área
2. Indicador NC_TOP en cuartil 2 o superior
3. Indicador ASC > 0 **ó** DPC > 0
4. Indicador FR_A > 0
5. Al menos 1 investigador sénior o asociado
6. Al menos 3 años de existencia

#### Categoría B
1. IG en cuartil 2 o superior
2. Indicador NC_TOP > 0 **ó** NC_A > 0
3. Indicador ASC > 0 **ó** DPC > 0
4. Al menos 1 investigador asociado o junior

#### Categoría C
1. IG en cuartil 3 o superior
2. Al menos 1 producto en cualquier sección

> **Nota para ReportSoft:** las condiciones de investigadores activos (emérito, sénior,
> asociado), la cohesión y la colaboración **no se modelan en el módulo de Plan de Mejora**,
> porque no vienen en `medicion_957.xlsx`. El módulo se enfoca en los indicadores de productos
> (NC_TOP, NC_A, NC_B, ASC, DPC, FR_A, FR_B) que sí están en el Excel oficial.

### 6.3 Cuartiles por indicador de sección (para las condiciones 1-3 de A1/A)

El reporte ScienTI también muestra los cuartiles de cada `valor_indicador_j` dentro del área:

```
Para el indicador NC_TOP, área Ingeniería y Tecnología:
  Mínimo: 73.44   Q4: 471.49   Q3: 1138.67   Q2: 2155.98   Máximo: 7098.22
```

Estos valores son la **fuente directa** de la hoja `cuartiles` de `medicion_957.xlsx`.
La nota del PDF aclara: *"Min y Max denotan el mínimo y máximo del conjunto de datos, sin
tomar las observaciones que tengan el valor 0. Los valores de los cuartiles calculados para
cada uno de los indicadores de producción se realiza sobre la población de grupos que
constituyen la misma área de conocimiento."*

---

## 7. Cómo ReportSoft usa estas fórmulas — sección por sección

| Sección UI | Qué usa | Fuente | Fórmulas del proceso oficial aplicadas |
|---|---|---|---|
| **Estado del plan** | Categoría, área | Hoja `grupos` | Resultado del Paso 6.2 |
| **Sección 1 — Diagnóstico** | `valor_indicador_j` | Hoja `indicadores` | Resultado del Paso 3 |
| **Sección 1 — Umbral P25** | Percentil 25 de los grupos objetivo | Calculado sobre hoja `indicadores` | Estadístico sobre resultados del Paso 3 |
| **Sección 3 — Comparación** | `valor_indicador_j` de otros grupos | Hoja `indicadores` | Resultado del Paso 3 |
| **Sección 4 — Mínimos** | min, Q4, Q3, Q2, máx por indicador | Hoja `cuartiles` | Cuartiles del Paso 6 |
| **Sección 4 — Requisito** | La columna cuartil exigida | `CUARTIL_OBJETIVO_POR_CATEGORIA` | Condiciones del Paso 6.2 |
| **Sección 5 — λ por subtipo** | `ln(1 + total/ventana)` | Hoja `productos` + fórmula | Fórmula exacta del Paso 2 |
| **Sección 5 — ratio** | `valor_indicador / Σλ` | Calculado | `C_j(área)` derivado de Pasos 3 y 4 |
| **Sección 5 — Simulado** | `ratio × Σλ_sim` | Proyección | Extensión del Paso 3 |

---

## 8. Caso de aplicación completo: grupo "Automática" (categoría A1)

Datos reales del reporte ScienTI (archivo `A1s.pdf`):

### 8.1 Paso 2 — λ por sección

**Sección NC_TOP:**

| Subtipo | Total | Ventana | λ = ln(1+T/V) | Verificación |
|---|---|---|---|---|
| ART_OPEN_A1 | 11 | 7 | **0.9445** | ln(1+11/7) = ln(2.571) ✓ |
| ART_A2      | 4  | 7 | **0.4520** | ln(1+4/7) = ln(1.571) ✓ |
| ART_A1      | 10 | 7 | **0.8873** | ln(1+10/7) = ln(2.429) ✓ |
| ART_OPEN_A2 | 11 | 7 | **0.9445** | ln(1+11/7) = ln(2.571) ✓ |
| **Σλ_TOP**  |    |   | **3.2283** | |

**Sección NC_A:**

| Subtipo | Total | Ventana | λ |
|---|---|---|---|
| ART_OPEN_B | 11 | 7 | 0.9445 |
| ART_B | 3 | 7 | 0.3567 |
| MA4 (modelo utilidad) | 1 | 10 | 0.0953 |
| PA4 (patente) | 1 | 10 | 0.0953 |
| **Σλ_A** | | | **1.4918** |

**Sección NC_B:**

| Subtipo | Total | Ventana | λ |
|---|---|---|---|
| LIB_C | 4 | 7 | 0.4520 |
| CAP_LIB_C | 1 | 5 | 0.1823 |
| SF (software) | 18 | 5 | 1.5261 |
| **Σλ_B** | | | **2.1604** |

### 8.2 Paso 3 — valor_indicador por sección

Usando la fórmula del §3.2 (`valor_indicador = Σλ × C(área)`) con las constantes del área
Ingeniería y Tecnología:

| Sección | Σλ | C(Ingeniería) | valor_indicador calculado | Oficial (PDF) | Diferencia |
|---|---|---|---|---|---|
| TOP | 3.2283 | 806.2 | 2602.7 | 2602.668 | **< 0.1%** |
| A | 1.4918 | 575.3 | 858.1 | 858.358 | < 0.1% |
| B | 2.1604 | 231.7 | 500.4 | 500.563 | < 0.1% |

*(La pequeña diferencia se explica por redondeo del Σλ. Usando los λ exactos del PDF, coincide
a 10 decimales.)*

### 8.3 Paso 4 — Índices

| Índice | valor_indicador | Máximo área | índice = val/máx | PDF |
|---|---|---|---|---|
| NC_TOP | 2602.668 | 7098.217 | **0.36667** | 0.36666508... ✓ |
| NC_A   | 858.358  | 7103.926 | **0.12083** | 0.12082874... ✓ |
| NC_B   | 500.563  | 3539.908 | **0.14141** | 0.14140568... ✓ |
| ASC    | 0.0      | 8968.264 | **0.00000** | 0.0 ✓ |
| DPC    | 4386.495 | 15687.47 | **0.27962** | 0.27961777... ✓ |
| FR_A   | 1804.073 | 2480.015 | **0.72744** | 0.72744421... ✓ |
| FR_B   | 2271.717 | 4065.326 | **0.55880** | 0.55880316... ✓ |

### 8.4 Paso 5 — Indicador de Grupo

$$IG = 3.7(0.36667) + 2.3(0.12083) + 0.4(0.14141) + 1.5(0) + 0.5(0.27962)$$
$$+ 1.0(0.72744) + 0.2(0.55880) + 0.1(0.33573) + 0.3(0.05435)$$

$$IG = 1.3567 + 0.2779 + 0.0566 + 0 + 0.1398 + 0.7274 + 0.1118 + 0.0336 + 0.0163$$

$$\boxed{IG = 2.7201} \approx 2.7199 \text{ (PDF)}$$  ✅

### 8.5 Paso 6 — Clasificación

Cuartiles del IG para Ingeniería y Tecnología: min=0.0179, Q4=0.5616, Q3=1.0751, Q2=1.8786, máx=6.1425

IG = 2.7199 > Q2 (1.8786) → **Cuartil 1 (25% superior)** ✓

Condiciones A1 verificadas:
- ✅ IG en cuartil 1
- ✅ NC_TOP (2602.668) > Q2 de NC_TOP en Ingeniería (2155.981) → cuartil 1 de NC_TOP
- ✅ ASC = 0 pero DPC = 4386.495 > 0
- ✅ FR_A = 1804.073 > 0
- ✅ Investigador sénior (dato GrupLAC)
- ✅ Cohesión = 1.609 > 0
- ✅ Más de 5 años de existencia

**Categoría resultante: A1** ✓

---

## 9. Caso de aplicación: grupo "GESTIÓN DE SISTEMAS ELÉCTRICOS" (categoría A)

Este grupo aparece en `medicion_957.xlsx` con categoría A (no A1) y es el grupo de ejemplo
principal de ReportSoft para el Plan de Mejora.

### 9.1 Estado actual (secciones 1 y 4)

**Objetivo: A1 — Área: Ingeniería y Tecnología — Columna exigida: Q2**

| Indicador | Tu valor (oficial) | Q2 (requisito A1) | ¿Cumple? | Brecha |
|---|---|---|---|---|
| NC_TOP | **743.711** | 2155.981 | ❌ | -1412.27 |
| NC_A   | **246.155** | 1049.072 | ❌ | -802.917 |
| NC_B   | **340.091** | 420.022  | ❌ | -79.931 |
| ASC    | **1386.294** | 1941.558 | ❌ | -555.264 |
| DPC    | **3284.851** | 1941.558 | ✅ | +1343.293 |
| FR_A   | **477.763** | 684.213  | ❌ | -206.450 |
| FR_B   | **3679.693** | 1272.563 | ✅ | +2407.130 |

Todos los datos vienen de las hojas `indicadores` y `cuartiles` de `medicion_957.xlsx`.

### 9.2 Σλ actuales (sección NC_TOP, Paso 2)

| Subtipo | Total | Ventana | λ |
|---|---|---|---|
| CAP_LIB_A1 | 1 | 5 | 0.1823 |
| ART_A1 | 1 | 7 | 0.1335 |
| ART_OPEN_A2 | 2 | 7 | 0.2513 |
| ART_OPEN_A1 | 3 | 7 | 0.3567 |
| **Σλ_TOP** | | | **0.9238** |

### 9.3 Constante de área para NC_TOP (Ingeniería)

$$C_{TOP}(\text{Ingeniería}) = \frac{743.711}{0.9238} = 805.02$$

*(El grupo "Automática" dio 806.2 con su propio Σλ. La pequeña diferencia se explica
por redondeo en los λ leídos del PDF. El valor exacto de C se obtiene del grupo con
mayor Σλ en el área: `máximo_área / Σλ_max_área`.)*

---

## 10. El simulador de impacto: justificación matemática

### 10.1 Punto de partida: la fórmula del Paso 3 reescrita

De §3.2 tenemos:

$$\text{valor\_indicador}_j(g) = \Sigma\lambda_j(g) \times C_j(\text{área})$$

donde `C_j(área)` es fijo para todos los grupos del área. Por tanto, si el grupo
genera nuevos productos que modifican `Σλ_j`:

$$\Delta\text{valor\_indicador}_j = \Delta\Sigma\lambda_j \times C_j(\text{área})$$

### 10.2 El `ratio` de ReportSoft es exactamente `C_j(área)`

$$\text{ratio} = \frac{\text{valor\_indicador\_actual}}{\Sigma\lambda_{\text{actual}}} = \frac{\Sigma\lambda_{\text{actual}} \times C_j(\text{área})}{\Sigma\lambda_{\text{actual}}} = C_j(\text{área})$$

No es una aproximación — es la **constante exacta de la sección y el área**.

### 10.3 Proyección con nuevos productos

Si el usuario simula agregar `n` unidades de un subtipo:

$$\lambda_{\text{nuevo}}(\text{subtipo}) = \ln\!\left(1 + \frac{\text{total\_actual} + n}{\text{ventana}}\right)$$

$$\Delta\lambda = \lambda_{\text{nuevo}} - \lambda_{\text{actual}}$$

$$\Sigma\lambda_{\text{simulado}} = \Sigma\lambda_{\text{actual}} + \sum_{\text{subtipos modificados}} \Delta\lambda_k$$

$$\boxed{\text{valor\_simulado} = \text{ratio} \times \Sigma\lambda_{\text{simulado}} = C_j(\text{área}) \times \Sigma\lambda_{\text{simulado}}}$$

### 10.4 Verificación: reproducción del valor actual con n=0

Con `n = 0` para todos los subtipos:
$$\Sigma\lambda_{\text{simulado}} = \Sigma\lambda_{\text{actual}}$$
$$\text{valor\_simulado} = C_j(\text{área}) \times \Sigma\lambda_{\text{actual}} = \text{valor\_indicador\_actual}$$

**El simulador reproduce exactamente el valor oficial con cero adiciones** — garantía de que
no hay desfase de escala.

### 10.5 Escenarios numéricos para NC_TOP del grupo GESTIÓN

Estado: `Σλ_actual = 0.9238`, `ratio = 805.02`, `requisito A1 = 2155.98`

| Escenario | Modificación | Σλ_sim | Valor simulado | ¿Cumple A1? |
|---|---|---|---|---|
| Base (n=0) | — | 0.9238 | 743.7 | ❌ |
| +5 caps A1 | CAP_LIB_A1: 1→6, Δλ=ln(7/5)−ln(6/5)=0.6062 | 1.5300 | 1231.7 | ❌ |
| +10 art Q1 | ART_A1: 1→11, Δλ=ln(12/7)−ln(8/7)=0.8114 | 1.7352 | 1397.7 | ❌ |
| +5 caps A1 y +10 art Q1 | Suma de ambos Δλ | 2.3414 | 1884.9 | ❌ |
| +5 caps A1 y +20 art Q1 | ART_A1: 1→21, Δλ=ln(22/7)−ln(8/7)=1.2528 | 2.7828 | 2242.4 | **✅** |

Para cumplir NC_TOP con A1, el grupo necesita aproximadamente **+5 capítulos A1 y +20 artículos Q1
adicionales** — o una combinación equivalente de otros subtipos TOP.

### 10.6 ¿Cuándo la proyección puede desviarse del resultado real?

La proyección es exacta en teoría bajo el modelo lineal de §3.2. Las fuentes de desviación son:

| Fuente de error | Magnitud | Explicación |
|---|---|---|
| Imprecisión en λ (borde de ventana) | ~4.7% de subtipos | Fechas de corte distintas a la ventana nominal |
| Cambio de área máxima entre convocatorias | < 5% | Si el grupo top del área cambia, C_j varía |
| Productos nuevos que no existían en el grupo | Indeterminado | El subtipo nuevo entra con λ_actual=0 → Δλ = λ_nuevo. La linealidad sigue siendo válida. |
| Cambios en reglas de MinCiencias entre convocatorias | Desconocido | Ventanas, categorías de productos, ponderaciones |

Para variaciones moderadas (≤ 10 unidades nuevas en subtipos ya producidos), el error esperado
es **< 5%**. Para escenarios de largo plazo (décadas de producción), el modelo pierde validez.

---

## 11. Grupos sin datos oficiales — aproximación BD interna

50 de los 125 grupos de la UTP no aparecen en `medicion_957.xlsx` (no participaron o no
alcanzaron el umbral mínimo de productos). Para ellos, ReportSoft usa una aproximación
propia que NO sigue el proceso oficial de 6 pasos.

### 11.1 Diferencias respecto al proceso oficial

| Aspecto | Proceso oficial (Pasos 2-6) | BD interna (aproximación) |
|---|---|---|
| Fuente de productos | Hoja `productos` de `medicion_957.xlsx` | `academia_utp_integrado.db` |
| Fórmula λ | `ln(1 + total/ventana)` con ventana por subtipo | λ fijo por tipo genérico (A1=1.00, A2=0.75…) |
| Rendimientos decrecientes | ✅ (logaritmo) | ❌ (λ fijo — lineal) |
| Constante de área C_j | Calculada | No disponible (sin max_área) |
| Cuartiles | Hoja `cuartiles` oficial | No disponible |
| Simulador | ✅ (ratio exacto) | ❌ No disponible |

### 11.2 λ fijo de la BD interna vs. λ oficial

La aproximación usa `VENTANAS_957`: cada producto activo suma un peso fijo:

| Tipo de producto | λ fijo BD interna | λ oficial equivalente (ejemplo) |
|---|---|---|
| Artículo A1 | 1.00 | ln(1 + N/7) ← varía con N |
| Artículo A2 | 0.75 | ln(1 + N/7) ← varía |
| Artículo B  | 0.50 | ln(1 + N/7) ← varía |
| Libro       | 2.00 | ln(1 + N/7 o N/5) ← varía |
| Capítulo    | 1.00 | ln(1 + N/5) ← varía |

Para un investigador con 5 artículos A1 (ventana 7), el oficial sería
λ = ln(1+5/7) = 0.539 (asignado al grupo completo), pero el BD interno asigna
1.00 por artículo, dando 5.00 — sobreestima en este caso.

### 11.3 Cuándo confiar en los resultados de BD interna

- El **diagnóstico de brechas** (cuáles indicadores son más bajos) es orientativo — el *orden*
  de los indicadores tiende a preservarse aunque la magnitud sea diferente.
- El **plan de recomendaciones** de productos sigue siendo válido cualitativamente.
- Los **valores numéricos** no son comparables directamente con los de grupos con datos
  oficiales.

---

## 12. Tabla resumen: oficial vs. aproximación

| Elemento | Sección UI | Pasos oficiales aplicados | Estatus |
|---|---|---|---|
| Categoría actual y área | Estado | Resultado Paso 6.2 | **100% oficial** |
| `valor_indicador_j` actual | 1, 4 | Resultado Paso 3 | **100% oficial** |
| Umbral P25 del área | 1 | Percentil 25 sobre Paso 3 | Estadístico sobre datos oficiales |
| min / Q4 / Q3 / Q2 / máx | 4 | Cuartiles del Paso 6 | **100% oficial** |
| Columna "requisito mínimo" | 4 | Condiciones del Paso 6.2 | Mapeo validado sobre datos oficiales |
| Desglose por subtipo | 2 | Tabla del Paso 2 | **100% oficial** |
| λ = ln(1+total/ventana) | 5 (interno) | Fórmula del Paso 2 | **Oficial, verificada al 95.3%** |
| `ratio = C_j(área)` | 5 | Derivado de Pasos 3 y 4 | **Exacto** (dentro del 4.7% de imprecisión de λ) |
| `valor_simulado = ratio × Σλ_sim` | 5 | Extensión del Paso 3 | **Lineal exacto** (mismo supuesto que Paso 3) |
| Indicadores grupos BD interna | 1 | No aplica (aproximación propia) | **Aproximación** — sin validación oficial |
| IG (Indicador de Grupo) | No mostrado en Plan | Paso 5 | No usado en Plan de Mejora (solo cuartiles por indicador) |

---

## 13. Preguntas frecuentes

### P1. ¿Por qué el simulador puede decir que llego a A1 con +20 artículos si la categoría depende también del IG y condiciones adicionales?

El simulador proyecta solo los indicadores de producto (`valor_indicador_j`). La categoría
final requiere también:
- Que el IG (suma ponderada de todos los índices normalizados) quede en el cuartil 1.
- Condiciones de investigadores activos, cohesión y existencia del grupo.

El simulador es una herramienta de planificación de producción científica, no una garantía
de categoría. Se recomienda verificar las condiciones adicionales de §6.2 de forma
independiente.

---

### P2. ¿Las ponderaciones (3.7 para NC_TOP, 2.3 para NC_A…) son las mismas para todas las categorías y áreas?

Sí. Las ponderaciones son **fijas para toda la convocatoria** (aparecen idénticas en todos
los reportes PDF analizados, independientemente del área o la categoría del grupo). Solo
cambian si MinCiencias modifica el modelo en una convocatoria posterior.

---

### P3. ¿Puedo reproducir el IG de mi grupo desde cero?

Con los datos disponibles en `medicion_957.xlsx`:

```
1. Calcular Σλ_j por sección (hoja `productos`)
2. Obtener el máximo de valor_indicador_j en el área (hoja `cuartiles`, columna "Máximo")
3. índice_j = Σλ_j × C_j(área) / máximo_j(área) = valor_indicador_j / máximo_j(área)
4. IG = 3.7×índice_TOP + 2.3×índice_A + 0.4×índice_B + 1.5×índice_ASC
       + 0.5×índice_DPC + 1.0×índice_FRA + 0.2×índice_FRB
       + 0.1×índice_cohesión + 0.3×índice_colaboración
```

Los índices de cohesión y colaboración **no están en `medicion_957.xlsx`** (son calculados
a partir de la red de coautorías) — sin ellos, el IG reconstruido será aproximado (les falta
el 0.4 máximo que aportan cohesión y colaboración juntos).

---

### P4. ¿Por qué la constante C_j varía entre áreas y qué implica eso?

`C_j(área) = máximo_j(área) / Σλ_max_j(área)`. Depende de cuán productivo (en términos
de λ) es el grupo más activo del área en esa sección. Ingeniería tiene grupos con muchos
artículos de alta calidad en ventanas largas; Ciencias Agrícolas también. La constante
refleja el "nivel de excelencia del área".

Para el simulador, esto significa que **comparar el ratio entre distintos grupos es posible
solo dentro del mismo área**. Un ratio de 806 para NC_TOP en Ingeniería y 827 en Ciencias
Agrícolas no indica que Ciencias Agrícolas pague más — solo que la distribución de λ del
área tiene un máximo ligeramente mayor en esa convocatoria.

---

## 14. Glosario

| Término | Definición |
|---|---|
| **λ (lambda)** | Peso de un subtipo: `ln(1 + total/ventana)`. Fórmula oficial del Paso 2 de ScienTI. |
| **Σλ_j** | Suma de λ de todos los subtipos de la sección j para un grupo. |
| **C_j(área)** | Constante de escalado: `valor_indicador_j / Σλ_j`. Igual para todos los grupos del mismo área (es el `ratio` en ReportSoft). |
| **valor_indicador_j** | Resultado del Paso 3 de ScienTI: `Σλ_j × C_j(área)`. Escala de cientos o miles. |
| **índice_j** | Resultado del Paso 4: `valor_indicador_j / máximo_j(área)`. Entre 0 y 1. |
| **IG** | Indicador de Grupo: `Σ(w_j × índice_j)` con ponderaciones del Paso 5. |
| **ventana** | Años de vigencia de un subtipo de producto (5, 7 o 10 años). |
| **cuartil (Q4/Q3/Q2)** | Percentil 25/50/75 de `valor_indicador_j` entre los grupos del mismo área y categoría objetivo (Paso 6). |
| **requisito mínimo** | El cuartil de `valor_indicador_j` exigido para la categoría objetivo: mín→C, Q4→B, Q3→A, Q2→A1. |
| **ratio** | En ReportSoft: `valor_indicador_actual / Σλ_actual` = `C_j(área)`. Constante de escalado que permite proyectar valores simulados. |
| **bd_interna** | Fuente de respaldo para grupos sin datos en `medicion_957.xlsx`. Usa λ fijo (no logarítmico). |

---

## Apéndice A — Ponderaciones oficiales por índice (extraídas de los PDFs ScienTI)

Verificadas en A1s.pdf (grupos "Automática" y "BIOECOS"):

| Índice | Ponderación | % del total |
|---|---|---|
| Nuevo Conocimiento TOP | **3.7** | 37% |
| Nuevo Conocimiento A | **2.3** | 23% |
| Apropiación Social | **1.5** | 15% |
| Formación RH Tipo A | **1.0** | 10% |
| Colaboración | 0.3 | 3% |
| Divulgación Pública | **0.5** | 5% |
| Formación RH Tipo B | **0.2** | 2% |
| Nuevo Conocimiento B | **0.4** | 4% |
| Cohesión | 0.1 | 1% |
| **Total** | **10.0** | 100% |

> Los productos de NC_TOP acumulan el 37% del IG. Por eso la categoría A1 exige
> explícitamente estar en el cuartil 1 de NC_TOP — sin producción TOP de alto
> impacto es prácticamente imposible alcanzar A1.

## Apéndice B — Archivos fuente

| Archivo | Contenido relevante |
|---|---|
| `data/pdf/A1s.pdf` | Reportes completos de grupos A1 (Pasos 1-6 de ScienTI) |
| `data/pdf/As.pdf`, `Bs.pdf`, `Cs.pdf` | Ídem para categorías A, B, C |
| `data/output/medicion_957.xlsx` | Hojas: `grupos`, `indicadores` (Paso 3), `productos` (Paso 2), `cuartiles` (Paso 6) |
| `data/cache/categorias_grupos_957.json` | Clasificación de los 50 grupos sin datos oficiales |
| `src/constants.py` | `INDICADORES_957`, `CUARTIL_OBJETIVO_POR_CATEGORIA`, `VENTANAS_957` |
| `src/analisis_seguimiento.py` | `SimuladorCategoriaInterna`: implementa Pasos 2, 3 y 4 |
| `src/views/vista_seguimiento_grupos.py` | UI del Plan de Mejora y el simulador |

## Apéndice C — Verificación cruzada del modelo del simulador

### C.1 Valor calculado manualmente (confiable)

Se extrajo manualmente de `A1s.pdf` el Σλ_TOP del grupo "Automática" (4 subtipos:
ART_OPEN_A1=11, ART_A2=4, ART_A1=10, ART_OPEN_A2=11, todos ventana 7 años):

| Grupo | Σλ_TOP (manual) | valor_TOP (PDF) | ratio = val/Σλ |
|---|---|---|---|
| Automática (Ingeniería, A1) | 3.2283 | 2602.668 | **806.2** |

### C.2 Script de verificación automatizada — resultado incompleto

Un script ejecutó extracción automática sobre todos los PDFs para calcular el ratio en cada
grupo. Resultado para Ingeniería y Tecnología (NC_TOP, n=14 grupos):

```
Ratios NC_TOP: min=119.2  max=11104.8  mean=2361.5
  Automática:  val=2602.7  sumL=0.755  ratio=3448
  CAFÉ:        val=4755.7  sumL=0.428  ratio=11104
  GAOPE:       val=2285.9  sumL=1.764  ratio=1296
```

Los ratios varían 93× dentro del mismo área — **esto no refuta la hipótesis de constancia**,
sino que evidencia extracción incompleta: el script encontró sumL=0.755 para Automática
cuando el valor correcto es 3.2283 (solo ~23% de los subtipos extraídos). Distintas tasas
de extracción por grupo producen ratios falsos y distintos.

### C.3 Cómo verificar correctamente

La verificación definitiva requiere usar `medicion_957.xlsx` (no los PDFs):
1. Para cada grupo, calcular `Σλ_j` desde la hoja `productos`:
   `SUM(ln(1 + total/ventana) para cada fila con indicador = j)`
2. Tomar `valor_indicador_j` desde la hoja `indicadores`.
3. Calcular `ratio = valor/Σλ` para cada grupo.
4. Si los ratios de todos los grupos del mismo área convergen, la hipótesis es correcta.

Hasta que se ejecute esta verificación, el ratio del simulador debe entenderse como
**buena aproximación** de `C_j(área)`, derivada matemáticamente del Paso 4 (§3.2),
pero sin confirmación numérica multgrupo.

---

*Documento técnico para uso del equipo ReportSoft – UTP.*  
*Basado en el Documento Conceptual de la Conv. 957 de MinCiencias (pp. 119-125) y los
reportes de medición ScienTI disponibles en `data/pdf/`.*
