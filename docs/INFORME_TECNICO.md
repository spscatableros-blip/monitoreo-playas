# Informe técnico del proyecto
## Tablero CONAGUA · Calidad del agua en sitios costeros

**Repositorio:** `github.com/spscatableros-blip/monitoreo-playas`
**Publicación:** GitHub Pages (sitio estático)
**Componentes:** tablero de playas · tablero de indicadores costeros · sistema automatizado de actualización de datos

---

## Índice

1. Objetivo y alcance del sistema
2. Cómo se construyó el repositorio (proceso y decisiones)
3. Arquitectura general
4. Estructura completa del repositorio
5. Las librerías: qué es cada una y qué hace exactamente
6. Los formatos de datos (`data.js`)
7. Los scripts de conversión, parte por parte
8. Cómo se transforman los datos (reglas detalladas)
9. Cómo se analizan los datos: fórmulas y algoritmos
10. Cómo se genera el tablero (render)
11. El robot de automatización (GitHub Actions)
12. Verificación y validación realizadas
13. Limitaciones conocidas y trabajo futuro
14. Glosario

---

# 1. Objetivo y alcance del sistema

El sistema publica y visualiza los resultados del **monitoreo de calidad del agua** en playas y
sitios costeros de México, a partir de los datos oficiales de **CONAGUA** y **COFEPRIS**.

Cubre dos conjuntos de información independientes:

| Tablero | Contenido | Cobertura |
|---|---|---|
| **Playas** | Clasificación sanitaria (APTA / NO APTA) por sitio de muestreo, según el criterio recreativo de enterococos fecales (≤200 NMP/100 mL) | 2014–2026 · 34 operativos · ~393 sitios · 641 playas · 17 estados |
| **Indicadores costeros** | Mediciones fisicoquímicas y bacteriológicas crudas (8 indicadores) | 2012–2025 · 23 447 muestras · 950 sitios · 17 estados |

Cada año tiene hasta **tres operativos de vigilancia** (temporadas): **Semana Santa**, **Verano** e
**Invierno**, que corresponden a los periodos vacacionales de mayor afluencia.

El objetivo técnico fue triple:

1. **Publicar** la información de forma pública, gratuita y sin infraestructura de servidor.
2. **Automatizar** la actualización, de modo que una persona sin conocimientos de programación
   pueda publicar datos nuevos subiendo un archivo Excel.
3. **Documentar** el sistema para permitir su mantenimiento y relevo.

---

# 2. Cómo se construyó el repositorio (proceso y decisiones)

## 2.1 Punto de partida

El repositorio existía ya con los dos tableros funcionando (`playas/` e `indicadores/`), pero la
actualización de datos era **manual**: había que editar a mano los archivos `data.js`, que contienen
miles de registros en una sola línea. El propio README lo advertía:

> *"La actualización es temporal: vive solo en tu navegador durante esa sesión. Para dejarla
> publicada de forma permanente hay que regenerar `playas/data.js` y hacer commit."*

El tablero permitía **cargar un Excel desde la página**, pero ese cambio existía solo en la sesión
del navegador: al recargar, desaparecía. No había forma práctica de publicar datos nuevos.

## 2.2 Diagnóstico

Se analizó el código existente y se determinó:

- El tablero de playas leía `window.DEFAULT_DATA` desde `playas/data.js`.
- El tablero de indicadores leía `window.IND_DATA` desde `indicadores/data.js`.
- **No existía ningún mecanismo para exportar el `data.js`**: los botones de descarga del tablero
  solo generaban `.xlsx`, `.png` y `.pdf`.
- La lógica de lectura de Excel del navegador estaba en dos funciones JavaScript:
  `prettyPeriod()` y `cleanSheet()` de `playas/index.html`.

**Decisión de diseño clave:** los conversores de Python debían **replicar exactamente** esas dos
funciones JavaScript. De ese modo, lo que el usuario ve al probar un archivo en la página es
idéntico a lo que queda publicado. Se evita así que existan dos interpretaciones distintas del
mismo archivo fuente.

## 2.3 Construcción por etapas

| Etapa | Qué se construyó | Resultado |
|---|---|---|
| 1 | `scripts/xlsx_a_datajs.py` — conversor de playas (formato por temporada) | Replica `cleanSheet()` y `prettyPeriod()`; mezcla incremental |
| 2 | `.github/workflows/actualizar-datos.yml` — el robot | Automatiza: subir Excel → regenerar → publicar |
| 3 | `scripts/xlsx_a_inddatajs.py` — conversor de indicadores | Validado byte-idéntico contra el `data.js` publicado |
| 4 | `scripts/semarnat_a_datajs.py` — conversor del reporte maestro SEMARNAT | Soporta el formato real de trabajo (una hoja por año) |
| 5 | Detección automática de formato en el robot | El usuario ya no elige script: el robot decide |
| 6 | `docs/` — documentación técnica | Permite mantenimiento y relevo |

## 2.4 Decisiones de diseño y su justificación

**a) Conversores independientes, no un solo script "universal".**
Cada formato de origen tiene una estructura muy distinta (por temporada, por año, base columnar).
Un único script con muchas ramas sería frágil y difícil de mantener. Se optó por tres scripts
autocontenidos, y que el **robot** decida cuál usar.

**b) Comportamiento distinto según el origen del dato.**

| Conversor | Comportamiento | Razón |
|---|---|---|
| `xlsx_a_datajs.py` | **Incremental** (mezcla) | El Excel puede traer solo una temporada nueva; el resto debe conservarse |
| `semarnat_a_datajs.py` | **Reemplazo total** | El archivo es el maestro histórico completo: es la fuente de verdad |
| `xlsx_a_inddatajs.py` | **Reemplazo total** | El Excel es un volcado íntegro de la base |

**c) Localización de columnas por nombre, no por posición.**
Los archivos oficiales cambian de ancho y orden entre años. Buscar `"Estado"`, `"Sitio"`,
`"Clasificación"` por su **texto normalizado** (sin acentos ni mayúsculas) hace el sistema tolerante
a esas variaciones. Un cambio de posición de columna no rompe nada.

**d) Validación por reconstrucción.**
Antes de dar por bueno un conversor, se **regeneró el `data.js` ya publicado** a partir del Excel
fuente y se comparó con el archivo original. Solo se aceptó el conversor cuando el resultado
coincidió. (Ver sección 12.)

**e) Cero dependencias externas en el navegador.**
Todas las librerías se sirven desde `lib/` dentro del repositorio; ninguna se carga de un CDN. El
tablero funciona sin conexión a internet y no depende de que un servicio de terceros siga
disponible en el futuro.

---

# 3. Arquitectura general

El sistema es un **sitio 100 % estático**: solo HTML, CSS y JavaScript. **No hay servidor de
aplicación ni base de datos.** Todo el procesamiento analítico ocurre en el navegador de quien
consulta el tablero.

```
   ETAPA 1: PREPARACIÓN          ETAPA 2: ANÁLISIS           ETAPA 3: PUBLICACIÓN
   (Python, en la nube)          (JavaScript, navegador)     (GitHub Actions + Pages)

   Excel  ──►  script  ──►  data.js  ──►  index.html  ──►  gráficas y tablas
                  ▲                                              ▲
                  └────── ejecutado por el robot ────────────────┘
```

El principio de diseño central es la **separación entre datos y presentación**:

- **`data.js`** contiene únicamente los datos, como una variable global de JavaScript.
- **`index.html`** contiene la lógica de análisis y el código de dibujo.

Actualizar los datos nunca requiere tocar el código del tablero, y viceversa.

---

# 4. Estructura completa del repositorio

```
monitoreo-playas/
│
├── index.html                   PÁGINA RAÍZ · pestañas + iframes (une los dos tableros)
├── README.md                    Documentación de uso
├── .gitignore                   Archivos excluidos del control de versiones
│
├── playas/                      ══ TABLERO 1: PLAYAS ══
│   ├── index.html               Lógica de análisis + render (829 líneas)
│   ├── data.js                  DATOS · window.DEFAULT_DATA        ← regenerado
│   ├── mexico.js                Geometría GeoJSON de los estados (mapa)
│   ├── assets/mapa.png          Mapa editorial de referencia
│   └── lib/
│       ├── chart.umd.min.js     Chart.js — motor de gráficas          (206 KB)
│       ├── annotation.min.js    Plugin de anotaciones                  (34 KB)
│       ├── datalabels.min.js    Plugin de etiquetas de datos           (13 KB)
│       └── xlsx.full.min.js     SheetJS — lectura de Excel            (882 KB)
│
├── indicadores/                 ══ TABLERO 2: INDICADORES COSTEROS ══
│   ├── index.html               Lógica de análisis + render
│   ├── data.js                  DATOS · window.IND_DATA             ← regenerado
│   └── lib/
│       ├── chart.umd.min.js     Chart.js                             (206 KB)
│       ├── datalabels.min.js    Plugin de etiquetas                   (13 KB)
│       └── boxplot.umd.min.js   Gráficas de caja (presente, no usado)  (19 KB)
│
├── datos/                       BUZÓN DE ENTRADA · Excel de playas
│   └── README.md                Instrucciones de formato
├── datos-indicadores/           BUZÓN DE ENTRADA · Excel de indicadores
│   └── README.md                Instrucciones de formato
│
├── scripts/                     ══ CONVERSORES (Excel → data.js) ══
│   ├── xlsx_a_datajs.py         Playas · formato por temporada
│   ├── semarnat_a_datajs.py     Playas · reporte maestro SEMARNAT
│   └── xlsx_a_inddatajs.py      Indicadores · base CONAGUA
│
├── .github/workflows/
│   └── actualizar-datos.yml     EL ROBOT · automatización completa
│
└── docs/
    ├── COMO_FUNCIONA.md         Guía técnica
    └── INFORME_TECNICO.md       Este informe
```

### Nota sobre `.gitignore`

Define qué archivos **no** deben subirse al repositorio, para mantenerlo limpio:

| Patrón | Excluye | Motivo |
|---|---|---|
| `.DS_Store` | Metadatos de carpetas de macOS | Basura del sistema operativo |
| `.claude/` | Configuración local de herramientas | Es local de cada persona |
| `*.bak`, `*.bak.*` | Respaldos (`data.js.bak.20260730_101500`) | Los genera el conversor al sobrescribir |
| `__pycache__/` | Bytecode compilado de Python | Se regenera automáticamente |

---

# 5. Las librerías: qué es cada una y qué hace exactamente

## 5.1 Cómo se cargan

Cada tablero declara sus dependencias con etiquetas `<script src="...">` de **ruta relativa**
apuntando a archivos **locales**. El navegador las descarga y ejecuta **secuencialmente, de arriba
hacia abajo**, antes de ejecutar el código propio del tablero.

`playas/index.html`, líneas 7–12:

```html
<script src="lib/chart.umd.min.js"></script>   <!-- 1 -->
<script src="lib/annotation.min.js"></script>  <!-- 2 -->
<script src="lib/datalabels.min.js"></script>  <!-- 3 -->
<script src="lib/xlsx.full.min.js"></script>   <!-- 4 -->
<script src="data.js"></script>                <!-- 5 -->
<script src="mexico.js"></script>              <!-- 6 -->
```

`indicadores/index.html`, líneas 7–9: solo `chart.umd.min.js`, `datalabels.min.js` y `data.js`.

## 5.2 Descripción de cada librería

### Chart.js (`chart.umd.min.js`, 206 KB)

Motor de graficación. Es la librería sobre la que se construyen **todas** las visualizaciones de
ambos tableros. Su funcionamiento:

- Recibe un elemento `<canvas>` del HTML y un objeto de configuración.
- Ese objeto define el **tipo** de gráfica (`bar`, `line`), los **datasets** (series de números),
  y las **opciones** (ejes, escalas, leyendas, interacción).
- Chart.js calcula la geometría (posición de barras, escalas, etiquetas de ejes) y **dibuja
  píxel a píxel** sobre el canvas.

Se emplean sus tipos `bar` (barras verticales y horizontales), `line` (líneas de tendencia), y
gráficas **mixtas** (barras + línea sobre dos ejes Y, como en la gráfica de resultados nacionales).

También se usa su **escala logarítmica** (`type:'logarithmic'`) en el tablero de indicadores para
enterococos y sólidos suspendidos, cuyos valores abarcan varios órdenes de magnitud (de 1 a más de
24 000 NMP/100 mL); en escala lineal, la mayoría de los puntos quedarían aplastados contra el eje.

**Debe cargarse primero**, porque crea el objeto global `Chart` del que dependen los plugins.

### chartjs-plugin-datalabels (`datalabels.min.js`, 13 KB)

Plugin que escribe **valores numéricos directamente sobre los elementos** de la gráfica (encima de
cada barra, junto a cada punto). Sin él habría que pasar el cursor por encima para leer un valor.

Se registra en Chart.js y se configura por dataset mediante la propiedad `datalabels`, que permite
definir posición (`anchor`, `align`), color, tipografía y un `formatter` (por ejemplo, añadir el
símbolo `%` o redondear decimales).

En el proyecto se desactiva globalmente por defecto (`Chart.defaults.plugins.datalabels.display =
false`) y se activa selectivamente solo donde aporta legibilidad.

### chartjs-plugin-annotation (`annotation.min.js`, 34 KB) — solo en playas

Plugin que superpone **elementos de referencia** sobre las gráficas: líneas, cajas, etiquetas.

En el tablero se usa para dibujar la **línea de promedio nacional** en la gráfica de cumplimiento
por estado: una línea vertical discontinua con su etiqueta, que permite ver de inmediato qué
estados están por debajo o por encima de la media.

### SheetJS / xlsx (`xlsx.full.min.js`, 882 KB) — solo en playas

Librería de lectura y escritura de archivos Excel **dentro del navegador**. Es la más pesada del
proyecto porque implementa el formato completo de Excel (XLSX es un contenedor ZIP con XML).

Cumple tres funciones:

1. **Lectura** (`XLSX.read`): cuando el usuario carga un `.xlsx` en la página, lo interpreta y lo
   convierte a una matriz de valores (`sheet_to_json` con `header:1`) que alimenta al tablero de
   forma temporal, sin publicar nada.
2. **Exportación de datos** (`XLSX.writeFile`): genera el archivo `playas_datos.xlsx` con los datos
   actualmente cargados, una hoja por operativo.
3. **Generación de plantilla**: crea `plantilla_playas.xlsx` con la estructura de ejemplo.

### `mexico.js` (146 KB) — datos, no librería

Contiene la **geometría GeoJSON** de las entidades federativas: la lista de coordenadas
(longitud, latitud) que forman el contorno de cada estado. Se usa para dibujar el mapa coroplético.

### `boxplot.umd.min.js` (19 KB) — presente pero no referenciado

Extensión de Chart.js para diagramas de caja. **El HTML actual no la carga**; es un archivo
disponible para un uso futuro (por ejemplo, mostrar la distribución de cada indicador por estado).

## 5.3 Por qué el orden de carga es obligatorio

```
chart.umd.min.js          crea el objeto global  Chart
        ↓
annotation + datalabels   se REGISTRAN dentro de Chart  →  requieren que ya exista
        ↓
data.js                   define  window.DEFAULT_DATA   →  los datos en memoria
        ↓
mexico.js                 define  MX_GEO                →  geometría del mapa
        ↓
<script> inline           USA todo lo anterior: lee los datos, calcula y llama a new Chart(...)
```

El registro explícito ocurre en las primeras líneas del código propio:

```js
if (window.ChartDataLabels) Chart.register(window.ChartDataLabels);
const _ann = window['chartjs-plugin-annotation'];
if (_ann) Chart.register(_ann.default || _ann);
```

Invertir el orden rompería el sistema: el código intentaría leer `window.DEFAULT_DATA` antes de que
`data.js` lo hubiera definido, o registrar plugins sobre un objeto `Chart` inexistente.

---

# 6. Los formatos de datos (`data.js`)

## 6.1 Playas — `playas/data.js`

Define una única variable global. Estructura de **lista de periodos con filas**:

```js
window.DEFAULT_DATA = {
  periods: [
    {
      key:   "verano24",           // identificador interno: temporada + año (2 dígitos)
      label: "Verano 2024",        // etiqueta visible en el tablero
      rows: [
        { est:"Baja California", dst:"Rosarito", playa:"Rosarito", sitio:"Rosarito I",  apta:1 },
        { est:"Baja California", dst:"Rosarito", playa:"Rosarito", sitio:"Rosarito II", apta:0 }
      ]
    }
    // … 34 objetos, uno por operativo, en orden cronológico
  ]
}
```

| Campo | Contenido | Tipo |
|---|---|---|
| `est` | Estado | texto |
| `dst` | Destino turístico | texto |
| `playa` | Playa | texto |
| `sitio` | Sitio de muestreo | texto |
| `apta` | Clasificación sanitaria | `1` = APTA · `0` = NO APTA · `null` = sin dato |

**Jerarquía:** Estado → Destino → Playa → Sitio. Una playa puede tener varios sitios de muestreo, y
esa relación es la base de la regla de agregación (sección 9.2).

**Formato físico:** el archivo es una sola línea de ~1.36 MB. No está pensado para lectura humana;
por eso los conversores son necesarios.

## 6.2 Indicadores — `indicadores/data.js`

Formato **columnar** (arreglos paralelos). El índice `i` identifica una muestra en todos los
arreglos simultáneamente:

```js
window.IND_DATA = {
  estados: ["BAJA CALIFORNIA", …],           // catálogo de 17 estados costeros
  est:  [13, 13, 15, …],                     // índice al catálogo · una entrada por muestra
  anio: [2012, 2012, 2012, …],               // año
  mes:  [10, 10, 10, …],                     // mes 1–12
  indicadores: ["ENTEROC_FEC", "SST", …],    // nombres de los 8 indicadores
  v: {                                       // valores, un arreglo por indicador
    ENTEROC_FEC: [3.0, 3.0, null, …],
    SST:         [28, 60, 54, …],
    …
  },
  meta: { muestras:23447, sitios:950, anio_min:2012, anio_max:2025,
          umbral_enteroc:200, fuente:"CONAGUA · BD calidad costeros" }
}
```

**Cómo se lee:** la muestra `i` corresponde al estado `estados[est[i]]`, tomada en `anio[i]`/`mes[i]`,
con enterococos `v.ENTEROC_FEC[i]` y sólidos `v.SST[i]`.

**Por qué columnar:** con 23 447 muestras, un arreglo de objetos repetiría el nombre de cada campo
23 447 veces. El formato columnar reduce drásticamente el tamaño y acelera los recorridos, que en
este tablero son mayoritariamente por columna (filtrar por año, calcular la mediana de un indicador).

### Los ocho indicadores

| Clave | Descripción | Unidad |
|---|---|---|
| `ENTEROC_FEC` | Enterococos fecales | NMP/100 mL |
| `SST` | Sólidos suspendidos totales | mg/L |
| `OD_%_SUP` | Oxígeno disuelto, % de saturación (superficie) | % |
| `OD_mg/L_SUP` | Oxígeno disuelto (superficie) | mg/L |
| `OD_%_MED` | Oxígeno disuelto, % de saturación (media agua) | % |
| `OD_mg/L_MED` | Oxígeno disuelto (media agua) | mg/L |
| `OD_%_FON` | Oxígeno disuelto, % de saturación (fondo) | % |
| `OD_mg/L_FON` | Oxígeno disuelto (fondo) | mg/L |

---

# 7. Los scripts de conversión, parte por parte

Los tres scripts comparten estructura: constantes → funciones auxiliares → funciones de parseo →
`construir()` → `main()` → bloque de arranque. Su única dependencia externa es **openpyxl**
(`pip install openpyxl`); todo lo demás es biblioteca estándar de Python.

| Módulo estándar | Uso en los scripts |
|---|---|
| `argparse` | Definir e interpretar las opciones de línea de comandos |
| `json` | Serializar el resultado al formato JSON que se incrusta en el `data.js` |
| `re` | Expresiones regulares: reconocer temporadas, años, nombres de hoja |
| `sys` | Terminar con mensaje de error (`sys.exit`) |
| `unicodedata` | Descomponer caracteres para eliminar acentos al normalizar |
| `pathlib.Path` | Manejo de rutas independiente del sistema operativo |
| `datetime` | Marca de tiempo de los respaldos; extracción del mes de una fecha |

## 7.1 `xlsx_a_datajs.py` — Playas, formato por temporada

**Entrada:** Excel donde **cada hoja es un operativo** (`verano24`, `semana.santa26`, `invierno2021`).
**Salida:** `playas/data.js`. **Modo:** incremental (mezcla).

### Funciones

**`norm(s)`** — Normaliza texto para comparaciones robustas:
```python
s = str(s).strip().lower()                     # sin espacios extremos, minúsculas
s = unicodedata.normalize("NFD", s)            # separa letra base y tilde
return "".join(c for c in s
               if unicodedata.category(c) != "Mn")   # elimina las tildes (marcas)
```
`"  Clasificación "` → `"clasificacion"`. Permite reconocer encabezados escritos con o sin acento,
en mayúsculas o minúsculas.

**`g(v)`** — Convierte el valor de una celda a texto recortado; devuelve `""` si es `None`. Evita
errores al operar sobre celdas vacías.

**`pretty_period(name)`** — Traduce el nombre de la hoja a etiqueta legible:
```python
m = re.match(r"^(semana\.?\s*santa|verano|invierno)\s*[._-]?\s*(\d{2}|\d{4})$", …)
```
El patrón acepta `semana.santa`, `semanasanta` o `semana santa`, seguido opcionalmente de un
separador (`.`, `_`, `-`) y un año de 2 o 4 dígitos. Si el año viene en 2 dígitos, antepone `"20"`.
`verano24` → `Verano 2024`. Si el nombre no coincide, aplica un embellecido genérico de respaldo.

**`clean_sheet(ws)`** — Convierte una hoja en lista de filas. Procedimiento:

1. Vuelca la hoja a matriz con `sheet_to_json`-equivalente (`iter_rows(values_only=True)`).
2. Busca en las **primeras 5 filas** una que contenga simultáneamente una columna `sitio…` y una
   `clasificacion…`; esa es la fila de encabezado.
3. Registra en un diccionario `col` el índice de columna de cada campo, comparando con `norm`:
   `estado` (exacto), `destino…` (prefijo), `playa` (exacto), `sitio…` y `clasificacion…` (prefijo).
4. Recorre las filas siguientes aplicando **arrastre** y la conversión de clasificación.

**`cargar_data_js(path)`** — Lee el `data.js` existente para poder mezclar. Recorta el prefijo
`window.DEFAULT_DATA = ` y el `;` final, y aplica `json.loads` sobre lo que queda.

**`escribir_data_js(path, data)`** — Serializa y escribe:
```python
payload = json.dumps(data, ensure_ascii=False, separators=(", ", ": "))
path.write_text(f"window.DEFAULT_DATA = {payload};", encoding="utf-8")
```
`ensure_ascii=False` conserva los acentos legibles; `separators` reproduce el estilo del archivo
original para que las diferencias en Git sean mínimas.

**`main()`** — Orquesta: interpreta opciones, recorre las hojas de todos los Excel indicados,
mezcla contra el `data.js` existente, imprime el resumen y escribe.

### Lógica de mezcla

```
Para cada temporada leída del Excel:
    ¿su clave ya existe en el data.js?
        SÍ  → se reemplaza esa temporada completa
        NO  → se agrega al final
Las temporadas no presentes en el Excel permanecen intactas.
```

Antes de sobrescribir crea un respaldo con marca de tiempo (`data.js.bak.AAAAMMDD_HHMMSS`), salvo
que se indique `--sin-respaldo` (opción usada en el robot, donde el respaldo no tiene sentido).

## 7.2 `semarnat_a_datajs.py` — Playas, reporte maestro SEMARNAT

**Entrada:** Excel maestro con **una hoja por año** (`2014`, `2024`…), en formato de reporte ancho.
**Salida:** `playas/data.js`. **Modo:** reemplazo completo.

### El formato de origen

Cada hoja-año tiene una estructura de dos niveles de encabezado:

```
Fila 0:  RELACIÓN MONITOREO DE PLAYAS - COFEPRIS · 2024                    (título)
Fila 1:  (vacía)
Fila 2:  No. | Estado | Destino turístico | Playa | Sitio | Coordenadas |
             Monitoreo prevacacional Verano 2024 | · | · |
             Monitoreo prevacacional Invierno 2024 | · | ·
Fila 3:  … | Calidad bacteriológica … Verano 2024 | · | · | …
Fila 4:  … | Latitud | Longitud | Fecha de muestreo | NMP/100 mL | Clasificación | …
Fila 5+: datos
```

Las columnas de identificación aparecen una sola vez; después vienen **bloques de 3 columnas por
temporada** (Fecha · NMP/100 mL · Clasificación). El número de bloques varía: los años completos
tienen 3 temporadas, 2024 solo 2.

### Constantes

```python
TEMP_KEY   = {"semana santa":"semana.santa", "verano":"verano", "invierno":"invierno"}
TEMP_LABEL = {"semana.santa":"Semana Santa", "verano":"Verano", "invierno":"Invierno"}
TEMP_RANK  = {"semana.santa":0, "verano":1, "invierno":2}
```

`TEMP_KEY` traduce del texto del encabezado a la clave interna; `TEMP_LABEL` de la clave a la
etiqueta visible; `TEMP_RANK` define el orden cronológico dentro de un año (Semana Santa ocurre en
marzo/abril, Verano en julio, Invierno en diciembre).

### `parse_hoja_anio(ws)` — el núcleo del parser

**Paso 1 — Localizar el encabezado.** Recorre las primeras 10 filas y toma como encabezado la
primera que contenga `"estado"` y alguna columna que empiece con `"sitio"`. Esto salta el título y
las filas vacías sin depender de posiciones fijas.

**Paso 2 — Mapear las columnas de identificación** (`estado`, `destino…`, `playa`, `sitio…`),
igual que en el conversor anterior.

**Paso 3 — Detectar los bloques de temporada.** Este es el mecanismo distintivo:

```python
for ci, c in enumerate(rows[hr]):
    m = re.search(r"(semana santa|verano|invierno)\s*(\d{4})", _norm(c))
    if m:
        bloques.append((ci, TEMP_KEY[m.group(1)], m.group(2)))
```

Recorre la fila de encabezado buscando el patrón *temporada + año de 4 dígitos*. Cada coincidencia
marca el inicio de un bloque y captura qué temporada y qué año es. Después, para cada bloque, busca
su columna de `Clasificación` dentro de sus límites (desde su inicio hasta el inicio del siguiente
bloque), usando la sub-fila de encabezados `hr+1`; si no la encuentra, asume la tercera columna del
bloque (`bc + 2`).

La clave se compone como `f"{temp}{yr[2:]}"`: `"verano"` + `"24"` = `verano24`.

**Paso 4 — Leer las filas de datos con arrastre.** Por cada fila con `sitio` no vacío, emite **una
fila por cada bloque de temporada** detectado, leyendo la clasificación de la columna
correspondiente a ese bloque.

### `construir(xlsx_path, incluir_2013)`

```python
wb = load_workbook(xlsx_path, read_only=True, data_only=True)
```
- `read_only=True`: modo de lectura eficiente en memoria, necesario para archivos grandes.
- `data_only=True`: lee el **valor calculado** de las celdas, no la fórmula.

Filtra las hojas con `re.fullmatch(r"\d{4}", nombre)`: solo procesa aquellas cuyo nombre es
exactamente un año de 4 dígitos, ignorando hojas auxiliares. Omite `2013` salvo `--incluir-2013`
(el propio reporte marca ese año como *"PRELIMINAR — pendiente de reconstrucción desde PDFs"*).

Finalmente ordena cronológicamente:
```python
orden = sorted(todos.values(), key=lambda i: (int(i["yr"]), TEMP_RANK[i["temp"]]))
```
Ordenamiento por **clave compuesta**: primero por año, y en caso de empate por el rango de la
temporada. Esto garantiza la secuencia SS → Verano → Invierno dentro de cada año.

## 7.3 `xlsx_a_inddatajs.py` — Indicadores costeros

**Entrada:** Excel de la base CONAGUA, hoja `indicadores`, una fila por muestra (23 columnas).
**Salida:** `indicadores/data.js`. **Modo:** reemplazo completo.

### Funciones

**`numerizar(x)`** — Convierte el valor de un indicador a número. Es la transformación más
específica del dominio:
```python
s = str(x).strip().replace("<", "").replace(">", "").replace(",", "")
return float(s)   # o None si no es convertible
```
Los laboratorios reportan **límites de detección** como texto: `<3` significa "por debajo del
límite de cuantificación de 3", y `>24196` "por encima del máximo medible". El criterio adoptado
—consistente con el `data.js` original— es **tomar el valor del límite**: `<3` → `3`,
`>24196` → `24196`. Celdas vacías o texto no numérico → `None`.

**`elegir_hoja(wb)`** — Devuelve la hoja cuyo nombre normalizado es `"indicadores"`; si no existe,
la primera del libro.

**`localizar_columnas(header)`** — Mapea por nombre: `CLAVE SITIO`, `FECHA REALIZACION` (prefijo,
tolera el acento), `AÑO`/`ANIO`, `ESTADO`, y los 8 indicadores comparados contra su forma
normalizada.

**`cargar_meta_previa(data_path)`** — Lee el `data.js` anterior para **conservar** los campos
`umbral_enteroc` y `fuente` de `meta`. Estos son parámetros de configuración, no datos derivados
del Excel; regenerar el archivo no debe perderlos.

**`construir(xlsx_path, meta_extra)`** — Recorre las filas y construye los arreglos columnares:

```python
idx = {_norm(s): i for i, s in enumerate(CATALOGO)}   # nombre normalizado → índice
for r in it:
    e = _norm(r[col["estado"]])
    if e not in idx:                    # estado fuera del catálogo costero
        descartados[…] += 1
        continue
    est.append(idx[e])
    anio.append(r[col["anio"]])
    f = r[col["fecha"]]
    mes.append(f.month if isinstance(f, datetime.datetime) else None)
    sitios.add(r[col["sitio"]])
    for k in INDICADORES:
        v[k].append(numerizar(r[col[k]]))
```

El catálogo se indexa **por nombre normalizado**, lo que resuelve automáticamente las variantes de
la fuente (`"COLIMA "` con espacio sobrante se normaliza a `"COLIMA"` y coincide).

Los metadatos derivados se calculan al final:
```python
meta = {"muestras": len(est),                    # filas válidas
        "sitios":   len(sitios),                 # CLAVE SITIO distintos
        "anio_min": min(anios_validos),
        "anio_max": max(anios_validos),
        **meta_extra}                            # umbral_enteroc y fuente heredados
```

## 7.4 El bloque de arranque

Los tres scripts terminan igual:

```python
if __name__ == "__main__":
    main()
```

Python asigna a la variable `__name__` el valor `"__main__"` cuando el archivo se ejecuta
**directamente**, y el nombre del módulo cuando se **importa**. Esta condición hace que `main()` se
ejecute solo en el primer caso.

**Consecuencia práctica:** los scripts son programas ejecutables, y **ninguna de sus funciones se
invoca desde otro archivo**. Aunque dos scripts tengan funciones con el mismo nombre (`construir`,
`_norm`), son funciones distintas e independientes en espacios de nombres separados.

Las cadenas de llamada internas son:

```
semarnat_a_datajs.py         xlsx_a_datajs.py             xlsx_a_inddatajs.py
  main()                       main()                       main()
   └─ construir()               ├─ clean_sheet()             ├─ cargar_meta_previa()
       └─ parse_hoja_anio()     ├─ pretty_period()           └─ construir()
           └─ _norm(), _g()     ├─ cargar_data_js()              ├─ elegir_hoja()
                                └─ escribir_data_js()            ├─ localizar_columnas()
                                                                 └─ numerizar()
```

---

# 8. Cómo se transforman los datos (reglas detalladas)

## 8.1 Localización de columnas por nombre

Todas las columnas se ubican comparando el encabezado **normalizado** (minúsculas, sin acentos, sin
espacios redundantes) contra un patrón. Unas se comparan por igualdad exacta (`estado`, `playa`) y
otras por prefijo (`destino…` capta *"Destino turístico"*; `sitio…` capta *"Sitio de muestreo"*;
`clasificacion…` capta *"Clasificación"*).

El sistema es por tanto **inmune** a: cambios de orden de columnas, columnas adicionales
intercaladas, diferencias de mayúsculas/acentos, y espacios sobrantes en los encabezados.

## 8.2 Arrastre de celdas combinadas

En los reportes oficiales, Estado, Destino y Playa aparecen **una sola vez** por grupo (celdas
combinadas verticalmente); las filas siguientes del mismo grupo tienen esas celdas vacías.

La regla implementada es: **si la celda tiene contenido, actualiza la variable; si está vacía,
conserva el último valor conocido.**

```python
if _g(r, col["est"]):   est = _g(r, col["est"])     # solo actualiza si hay contenido
…
rows.append({"est": est, …})                       # siempre escribe el valor vigente
```

| Excel | est | dst | playa | sitio |
|---|---|---|---|---|
| `BC \| Rosarito \| Rosarito \| Rosarito I` | BC | Rosarito | Rosarito | Rosarito I |
| `(vacío) \| (vacío) \| (vacío) \| Rosarito II` | BC ↓ | Rosarito ↓ | Rosarito ↓ | Rosarito II |

## 8.3 Criterio de fila válida

**`Sitio` es el campo obligatorio.** Una fila sin sitio se descarta (`continue`). Esto elimina
automáticamente filas separadoras, subtotales y espacios en blanco del reporte, sin necesidad de
reglas adicionales.

## 8.4 Traducción de la clasificación

```python
s = valor_celda.lower()
apta = 0 if "no apta" in s else (1 if "apta" in s else None)
```

El orden de evaluación es **crítico**: se comprueba `"no apta"` **antes** que `"apta"`, porque la
cadena `"no apta"` **contiene** la subcadena `"apta"`. Invertir el orden clasificaría erróneamente
todas las playas no aptas como aptas.

Es una comparación por **contenido**, no por igualdad, de modo que acepta variantes como
`"APTA"`, `"Apta "`, `"No Apta"`, `"NO APTA (2)"`. Cualquier otro texto o celda vacía → `None`
(sin dato), que se distingue explícitamente de "no apta".

## 8.5 Filtrado por catálogo de estados (indicadores)

Solo se conservan las muestras cuyo `ESTADO` normalizado pertenece al catálogo de **17 estados
costeros**. Las demás se descartan y se reportan en el resumen. En el archivo real esto elimina 2
de 23 449 filas: una de `SAN LUIS POTOSÍ` (estado sin litoral) y una con el estado vacío.

## 8.6 Derivación del mes

El Excel de indicadores no tiene columna de mes; se obtiene de la fecha de muestreo:
```python
mes.append(f.month if isinstance(f, datetime.datetime) else None)
```
La comprobación de tipo evita fallos cuando la celda contiene texto en lugar de una fecha real.

## 8.7 Serialización final

El resultado se convierte a JSON y se envuelve en la asignación que el navegador espera:

```
playas:       window.DEFAULT_DATA = { … };
indicadores:  window.IND_DATA     = { … };
```

Este envoltorio es lo que permite cargar los datos con una simple etiqueta `<script src="data.js">`,
sin necesidad de una petición HTTP asíncrona ni de un servidor: al ejecutarse el archivo, la
variable global queda definida.

---

# 9. Cómo se analizan los datos: fórmulas y algoritmos

**El análisis no lo realizan los scripts de Python.** Los conversores solo limpian y transforman.
Todo el cálculo analítico ocurre en **JavaScript, dentro del navegador**, cada vez que alguien abre
el tablero o cambia un filtro.

## 9.1 Criterio sanitario base

El umbral normativo aplicado es el **criterio recreativo de contacto primario**:

> **≤ 200 NMP/100 mL de enterococos fecales** → APTA
> **> 200 NMP/100 mL** → NO APTA

En el tablero de playas ese criterio ya viene aplicado en el campo `apta` del dato oficial. En el
tablero de indicadores se aplica directamente sobre el valor medido (`meta.umbral_enteroc = 200`).

## 9.2 Regla de agregación: cuándo una playa es APTA

Es la fórmula central del tablero de playas. Existen dos **niveles de medición** que el usuario
alterna con el selector superior:

**Nivel `sitio`** — cada muestra cuenta individualmente:

$$\text{\% cumplimiento} = \frac{\text{n.º de sitios con } apta = 1}{\text{n.º de sitios con dato}} \times 100$$

**Nivel `playa`** (predeterminado) — se agrupa por la clave `est|dst|playa` y se aplica el
**criterio conservador**:

$$\text{playa APTA} \iff \forall\, s \in \text{sitios(playa)} : apta_s = 1$$

Es decir, **basta con que un solo sitio resulte NO APTA para que toda la playa se contabilice como
NO APTA**. Implementación:

```js
function counts(rows){
  const g = {};
  for (const r of rows)
    (g[r.est+'|'+r.dst+'|'+r.playa] ??= []).push(r.apta);   // agrupa por playa

  let a = 0, t = 0;
  for (const k in g) {
    const v = g[k].filter(valid);        // valid(a) = (a===0 || a===1)
    if (!v.length) continue;             // playa sin ningún dato: no se cuenta
    t++;                                 // playa contable
    if (v.every(x => x === 1)) a++;      // APTA solo si TODOS sus sitios lo son
  }
  return { aptas:a, noaptas:t-a, total:t, pct: t ? a/t*100 : null };
}
```

**Tratamiento de los datos faltantes:** `valid()` excluye los `null` del cálculo. Una playa sin
ningún dato no entra en el denominador (no cuenta ni a favor ni en contra). Esto evita que la
ausencia de medición se interprete como incumplimiento.

## 9.3 Promedio anual

Cada año agrupa hasta tres operativos. El indicador anual es la **media aritmética de los
porcentajes de sus temporadas**, ignorando las que no tengan dato:

$$\overline{p}_{a\tilde{n}o} = \frac{1}{|T'|}\sum_{t \in T'} p_t \quad \text{donde } T' = \{t : p_t \neq \text{null}\}$$

```js
const avgNN = a => {
  const v = a.filter(x => x != null);
  return v.length ? v.reduce((x,y) => x+y, 0) / v.length : null;
};
```

Nótese que es un promedio **de porcentajes** (no ponderado por número de playas). Dado que el
universo de playas es prácticamente constante entre temporadas del mismo año, la diferencia
respecto de un promedio ponderado es marginal.

El año se extrae de la etiqueta del periodo mediante expresión regular:
```js
const periodYear = p => (String(p.label).match(/\d{4}/) || [''])[0];
```

## 9.4 Conteo de incumplimientos recurrentes

Identifica sitios con problemas sistemáticos. Recorre **todos** los operativos y acumula por sitio:

$$n_s = \sum_{t \in \text{operativos}} \mathbb{1}[apta_{s,t} = 0]$$

```js
function chronicSites(minN){
  const cnt = {};
  DATA.periods.forEach(p => p.rows.forEach(r => {
    const k = r.est+'|'+r.dst+'|'+r.playa+'|'+r.sitio;
    cnt[k] ??= { …, n:0 };
    if (r.apta === 0) cnt[k].n++;
  }));
  return Object.values(cnt).filter(x => x.n >= minN).sort((a,b) => b.n - a.n);
}
```

Umbral de reporte: **n ≥ 3**. Codificación cromática: **≥5** rojo (crónico), **4** naranja,
**3** dorado.

## 9.5 Ranking de destinos

Acumula incumplimientos por destino sobre toda la serie:

$$N_d = \sum_{t}\sum_{s \in d} \mathbb{1}[apta_{s,t} = 0]$$

Se ordena descendentemente y se muestran los N primeros (configurable: 10, 15 o 20).

## 9.6 Detección de tendencia (informe automático)

El informe en texto determina si el problema crece o disminuye comparando los **promedios de los
tres primeros y los tres últimos años** de la serie:

```js
const early = noByYear.slice(0,3), late = noByYear.slice(-3);
const eAvg = early.reduce((a,b)=>a+b,0) / early.length;
const lAvg = late.reduce((a,b)=>a+b,0)  / late.length;
const dir = lAvg > eAvg+0.5 ? 'al alza'
          : lAvg < eAvg-0.5 ? 'a la baja'
          : 'estable';
```

La **banda muerta de ±0.5** evita calificar como tendencia una fluctuación mínima. Es una
heurística descriptiva, no una prueba estadística de significancia.

También determina la **temporada crítica**, acumulando incumplimientos por tipo de temporada
(Semana Santa / Verano / Invierno) y tomando el máximo.

## 9.7 Escala cromática del mapa de calor

Los porcentajes se traducen a color mediante **interpolación lineal por tramos** sobre seis puntos
de control:

| % | Color |
|---|---|
| 50 | `#C0392B` rojo oscuro |
| 60 | `#E74C3C` rojo |
| 70 | `#E67E22` naranja |
| 80 | `#F1C40F` amarillo |
| 90 | `#ABEBC6` verde claro |
| 100 | `#27AE60` verde |

```js
v = Math.max(50, Math.min(100, v));         // recorte al rango [50,100]
// localizar el tramo [a,b] que contiene v y calcular la posición relativa:
const t = (v - a) / (b - a);
return mix(ca, cb, t);
```

La mezcla se hace canal por canal en RGB:

$$c_{\text{resultado}} = c_a + (c_b - c_a)\cdot t \quad \text{para cada canal } R,G,B$$

El recorte inferior en 50 % concentra la resolución cromática en el rango donde efectivamente se
encuentran los datos. Los valores nulos se pintan en gris (`#eef1f4`).

Adicionalmente, el color del texto se invierte a blanco cuando el fondo es oscuro (`v < 78`), para
mantener la legibilidad.

## 9.8 Proyección cartográfica del mapa

El mapa coroplético se genera dibujando SVG a partir del GeoJSON. La proyección aplicada es una
**equirrectangular con corrección de latitud media**:

```js
const k = Math.cos((minLat + maxLat) / 2 * Math.PI / 180);   // factor de corrección
const scale = Wt / ((maxLon - minLon) * k);                   // escala para ancho fijo (580 px)
const proj = (x, y) => [ (x - minLon) * k * scale,            // X
                         (maxLat - y) * scale ];              // Y (invertida)
```

- El factor **k = cos(latitud media)** compensa la convergencia de los meridianos: sin él, México
  aparecería estirado horizontalmente.
- La coordenada **Y se invierte** (`maxLat - y`) porque en SVG el eje vertical crece hacia abajo,
  mientras que la latitud crece hacia el norte.

Cada polígono se convierte en un atributo `d` de trazado SVG (`M` = mover, `L` = línea, `Z` =
cerrar), y se rellena con el color correspondiente a su porcentaje de cumplimiento.

## 9.9 Análisis del tablero de indicadores

Al trabajar con **valores continuos** (no clasificaciones binarias), este tablero emplea
estadística descriptiva.

### Mediana

```js
function median(a){
  const s = [...a].sort((x,y) => x-y);
  const m = s.length >> 1;                       // división entera por 2
  return s.length % 2 ? s[m] : (s[m-1] + s[m]) / 2;
}
```

**Se usa la mediana en lugar de la media** porque las distribuciones de enterococos y sólidos
suspendidos son **fuertemente asimétricas**: unos pocos valores extremos (>24 000 NMP) desplazarían
la media y darían una imagen distorsionada de la situación típica. La mediana es robusta frente a
esos valores atípicos.

### Cuantiles (interpolación lineal)

```js
function quantile(s, q){
  const p = (s.length-1) * q, b = Math.floor(p), r = p - b;
  return s[b+1] !== undefined ? s[b] + r*(s[b+1] - s[b]) : s[b];
}
```

Corresponde al **método 7** (el predeterminado en R y NumPy): posición $p = (n-1)q$, con
interpolación lineal entre los dos valores adyacentes.

Se emplea para fijar el rango cromático de los mapas de calor en los **percentiles 5 y 95** de las
medianas estado-año, en lugar del mínimo y el máximo absolutos. Así, un único valor extremo no
comprime toda la escala de color.

### Porcentaje de excedencia

$$\text{\% excedencia} = \frac{\#\{i : v_i > 200\}}{\#\{i : v_i \neq \text{null}\}} \times 100$$

```js
function pctExceed(idx){
  let t=0, ex=0;
  for (const i of idx){
    const e = D.v['ENTEROC_FEC'][i];
    if (e != null) { t++; if (e > 200) ex++; }
  }
  return t ? ex/t*100 : null;
}
```

Es el indicador principal de este tablero: qué proporción de las muestras superó el umbral
sanitario. Se calcula por estado, por año y por mes.

**Filtro de robustez:** en el ranking por estado se exigen **al menos 5 muestras** (`n >= 5`) para
que un estado aparezca. Sin ese filtro, un estado con 2 muestras y 1 excedencia figuraría con 50 %,
compitiendo indebidamente con estados que tienen cientos de muestras.

### Escala logarítmica

```js
const LOG_INDS = new Set(['ENTEROC_FEC','SST']);
…
type: LOG_INDS.has(indicador) ? 'logarithmic' : 'linear'
```

Enterococos y sólidos suspendidos se grafican en **escala logarítmica** porque su rango abarca
varios órdenes de magnitud (de 1 a >24 000). En escala lineal, el 90 % de las observaciones quedaría
comprimido contra el eje. Los indicadores de oxígeno disuelto, con rango acotado, usan escala lineal.

### Umbral de hipoxia

En el informe automático se calcula la proporción de muestras con **oxígeno disuelto < 4 mg/L**,
valor de referencia por debajo del cual se considera que hay estrés para la vida acuática.

## 9.10 Filtrado jerárquico

La función de filtrado es común a todo el tablero de playas:

```js
const sub = (rows, est, dst) =>
  rows.filter(r => (!est || r.est === est) && (!dst || r.dst === dst));
```

Los argumentos vacíos actúan como comodín: sin argumentos devuelve todo (nacional); con estado
filtra por estado; con ambos, por destino. Todas las visualizaciones se recalculan aplicando este
filtro, de modo que el mismo código sirve para la vista nacional y las vistas desagregadas.

---

# 10. Cómo se genera el tablero (render)

## 10.1 La página raíz

`index.html` de la raíz no contiene análisis ni librerías. Su función es integrar los dos tableros
mediante **iframes con carga diferida**:

```js
function show(target){
  const f = frames[target];
  if (!f.src) {                       // solo la primera vez que se activa la pestaña
    loading.style.display = 'flex';
    f.addEventListener('load', () => loading.style.display='none', {once:true});
    f.src = f.dataset.src;            // asigna el src → dispara la carga
  }
  …
}
```

El atributo real `src` se asigna **solo al abrir la pestaña por primera vez**. De este modo, quien
consulta únicamente el tablero de playas nunca descarga los ~1.4 MB de datos de indicadores.

## 10.2 Secuencia de arranque de un tablero

```
1. El navegador descarga y ejecuta las librerías de lib/   → objeto Chart disponible
2. Se registran los plugins                                → Chart.register(...)
3. Se ejecuta data.js                                      → window.DEFAULT_DATA definido
4. Se ejecuta mexico.js                                    → MX_GEO definido
5. Se ejecuta el <script> inline:
      let DATA = window.DEFAULT_DATA;
      …
      if (DATA && DATA.periods) renderAll();               → arranca el render
```

## 10.3 La función `renderAll()`

Punto único de reconstrucción del tablero completo:

```js
function renderAll(){
  renderKPIs(); fillReportScope(); renderReport(); renderAlert();
  renderG1(); renderG1b(); renderG2(); renderG3();
  renderHeat(); renderFigB(); fillExpStates(); renderExp(); renderMap();
  …
}
```

Cada `renderX()` es autónoma: recalcula sus propios agregados a partir de `DATA` y `LEVEL`, y
redibuja su gráfica. No hay estado intermedio compartido, lo que evita inconsistencias.

## 10.4 Ciclo de actualización

Antes de crear una gráfica se destruye la anterior:

```js
function destroy(id){
  if (charts[id]) { charts[id].destroy(); delete charts[id]; }
}
```

Chart.js registra escuchadores de eventos y animaciones por instancia; omitir `destroy()` provocaría
**fugas de memoria** y superposición de gráficas al cambiar de filtro.

Los eventos que disparan un nuevo render:

| Control | Efecto |
|---|---|
| Selector Playas / Muestras | Cambia `LEVEL` → `renderAll()` completo |
| Selector de estado (explorador) | `renderExp()` |
| Selector de alcance del informe | `renderReport()` |
| Selector de operativo (mapa) | `renderMap()` |
| Casilla "solo estados con incumplimiento" | `renderHeat()` |
| Carga de un `.xlsx` | Reemplaza `DATA` en memoria → `renderAll()` |

## 10.5 Exportación

- **Gráficas a PNG:** se copia el `<canvas>` a uno temporal con fondo blanco (el canvas original es
  transparente) y se descarga vía `toDataURL('image/png')`.
- **Mapa a PNG:** se serializa el SVG a texto, se codifica en base64 como imagen, se dibuja sobre un
  canvas al **doble de resolución** (`sc = 2`) y se descarga.
- **Datos a Excel:** se reconstruye un libro con una hoja por operativo usando SheetJS.
- **PDF:** se delega en la impresión del navegador; una regla `@media print` oculta los controles.

---

# 11. El robot de automatización (GitHub Actions)

Definido en `.github/workflows/actualizar-datos.yml`. Es el componente que convierte el proceso
manual en automático.

## 11.1 Disparadores

```yaml
on:
  push:
    paths:
      - "datos/**.xlsx"
      - "datos-indicadores/**.xlsx"
      - "scripts/*.py"
  workflow_dispatch:
```

Se ejecuta cuando se sube o modifica un Excel en cualquiera de los dos buzones, o cuando se cambia
un script (para reprocesar con la lógica nueva). `workflow_dispatch` permite además lanzarlo
manualmente desde la pestaña *Actions*.

## 11.2 Permisos y control de concurrencia

```yaml
permissions:
  contents: write        # necesario para que el robot pueda hacer commit

concurrency:
  group: actualizar-datos
  cancel-in-progress: false
```

El permiso de escritura debe habilitarse además en la configuración del repositorio
(*Settings → Actions → Workflow permissions → Read and write*). El grupo de concurrencia impide que
dos ejecuciones simultáneas intenten escribir a la vez; `cancel-in-progress: false` garantiza que
una ejecución en curso termine en lugar de abortarse.

## 11.3 Estructura de trabajos

Dos trabajos, **encadenados**:

```yaml
jobs:
  playas:        …
  indicadores:
    needs: playas          # espera a que termine playas
```

El encadenamiento evita que ambos intenten hacer `git push` simultáneamente sobre la misma rama, lo
que provocaría un rechazo por conflicto.

## 11.4 Pasos de cada trabajo

**1. Preparación del entorno**
```yaml
- uses: actions/checkout@v5        # descarga el repositorio
- uses: actions/setup-python@v6    # instala Python 3.12
- run: pip install openpyxl        # única dependencia
```

**2. Verificación de existencia**
```bash
if ls datos/*.xlsx >/dev/null 2>&1; then
  echo "hay=1" >> "$GITHUB_OUTPUT"
else
  echo "hay=0" >> "$GITHUB_OUTPUT"
fi
```
Si el buzón está vacío, los pasos siguientes se omiten (`if: steps.check.outputs.hay == '1'`) y el
trabajo termina correctamente sin hacer nada.

**3. Detección automática de formato** (solo en playas)
```bash
for f in datos/*.xlsx; do
  if python3 -c "…any(re.fullmatch(r'\d{4}', s) for s in wb.sheetnames)…" "$f"; then
    python3 scripts/semarnat_a_datajs.py "$f"          # reporte por año
  else
    python3 scripts/xlsx_a_datajs.py "$f" --sin-respaldo  # por temporada
  fi
done
```
Abre el Excel y examina los **nombres de sus hojas**: si alguna es un año de 4 dígitos, se trata del
reporte maestro SEMARNAT; en caso contrario, del formato por temporada. Quien sube el archivo no
necesita saber qué conversor corresponde.

En indicadores, si hay varios archivos se usa el último por orden alfabético:
```bash
ULTIMO=$(ls -1 datos-indicadores/*.xlsx | sort | tail -1)
```

**4. Commit condicional**
```bash
if git diff --quiet -- playas/data.js; then
  echo "playas/data.js no cambió; no se hace commit."
  exit 0
fi
git config user.name  "github-actions[bot]"
git config user.email "41898282+github-actions[bot]@users.noreply.github.com"
git add playas/data.js
git commit -m "Actualiza playas/data.js desde datos/ [skip ci]"
git push
```

Dos salvaguardas relevantes:
- **`git diff --quiet`**: solo se hace commit si el archivo realmente cambió. Volver a subir el
  mismo Excel no genera ruido en el historial (el proceso es **idempotente**).
- **`[skip ci]`**: instruye a GitHub Actions a no disparar un nuevo flujo con este commit, evitando
  un bucle infinito.

**5. Publicación** — GitHub Pages detecta el commit en `main` y republica el sitio automáticamente
(1–2 minutos). No requiere configuración adicional en el flujo.

## 11.5 Flujo completo

```
Usuario sube Excel a datos/ (o datos-indicadores/) desde github.com
        │
        ▼
GitHub Actions detecta el push sobre esa ruta
        │
        ▼
Instala Python + openpyxl · detecta el formato · ejecuta el conversor
        │
        ▼
¿Cambió el data.js?  ──no──►  fin (sin commit)
        │ sí
        ▼
commit + push como github-actions[bot]
        │
        ▼
GitHub Pages republica  →  tablero actualizado en 1–2 min
```

---

# 12. Verificación y validación realizadas

El criterio adoptado fue: **un conversor solo se considera correcto si reproduce el `data.js` ya
publicado a partir del Excel fuente.**

## 12.1 Conversor de indicadores

Se regeneró el archivo completo desde `indicadores-cost.xlsx` y se comparó campo por campo contra
el `indicadores/data.js` en producción:

| Comprobación | Resultado |
|---|---|
| Número de muestras | 23 447 = 23 447 ✔ |
| Número de sitios | 950 = 950 ✔ |
| Arreglo `est` | idéntico ✔ |
| Arreglo `anio` | idéntico ✔ |
| Arreglo `mes` | idéntico ✔ |
| Los 8 arreglos de `v` | idénticos ✔ |

La coincidencia exacta confirmó las reglas inferidas: el filtrado por catálogo costero
(23 449 → 23 447 filas; 952 → 950 sitios) y la numerización de límites de detección.

## 12.2 Conversor SEMARNAT

Se parseó el reporte maestro y se comparó contra las 33 temporadas publicadas:

| Resultado | Temporadas |
|---|---|
| Contenido idéntico | 27 |
| Diferencias reales del archivo maestro (correcciones de nomenclatura y sitios adicionales en 2022–2023) | 6 |
| Aportadas por el maestro y no publicadas aún | 1 (`verano26`) |

Las diferencias se inspeccionaron individualmente y correspondían a **correcciones legítimas del
archivo fuente** (`El Paraíso` → `Paraíso`, `Playa Uaymtún` → `Uaymitún`, cuatro sitios adicionales),
no a errores del conversor.

## 12.3 Pruebas funcionales adicionales

- **Idempotencia:** procesar dos veces el mismo Excel no produce cambios en el `data.js`
  (verificado con `git diff --quiet`).
- **Mezcla incremental:** con un Excel de prueba de dos hojas se comprobó que una temporada nueva se
  agrega y una existente se reemplaza, sin afectar a las demás (33 → 34 periodos).
- **Detección de formato:** se verificó que el reporte maestro se identifica como SEMARNAT y un
  archivo por temporada como formato por temporada.
- **Validez del resultado:** cada `data.js` generado se volvió a analizar sintácticamente
  (prefijo, JSON válido, sufijo `;`) y se comprobó el esquema de sus filas.
- **Casos límite manejados:** archivo de bloqueo de Excel (`~$…`) cuando el libro está abierto;
  nombres de archivo con espacio inicial; celdas de fecha con contenido no-fecha.

---

# 13. Limitaciones conocidas y trabajo futuro

## 13.1 Limitaciones

| Limitación | Descripción |
|---|---|
| **2013 preliminar** | El reporte marca ese año como pendiente de reconstrucción desde PDFs; se excluye por defecto |
| **2020 ausente** | No hay operativos registrados (contingencia sanitaria) |
| **Sin validación de contenido** | El robot publica lo que reciba: un Excel incompleto se publicaría sin advertencia |
| **Promedios no ponderados** | El % anual promedia porcentajes de temporada sin ponderar por número de playas |
| **Tendencia heurística** | La detección "al alza / a la baja" no es una prueba estadística de significancia |
| **Sin fecha de actualización visible** | El tablero no indica cuándo se actualizaron los datos por última vez |

## 13.2 Mejoras propuestas

**Prioridad alta**
1. **Validación previa a la publicación**: rechazar automáticamente una actualización que reduzca
   drásticamente el número de sitios de una temporada existente, o que introduzca nombres de estado
   no reconocidos.
2. **Sello de actualización**: añadir `meta.generado` al `data.js` y mostrarlo en el tablero.
3. **Pruebas automatizadas** de los conversores (con un Excel de muestra reducido en `tests/`),
   ejecutadas por el robot antes de publicar.

**Prioridad media**
4. **Registro de cambios**: que el robot escriba qué temporadas y sitios cambiaron en cada
   actualización.
5. **Ficha por playa**: historial individual completo de una playa, con detalle por sitio de
   muestreo *(implementada; pendiente de publicación)*.
6. **Estado compartible por URL**: reflejar los filtros activos en la dirección web.

**Prioridad baja**
7. Cruce analítico entre ambos tableros (clasificación de playas frente a valores medidos de
   enterococos en los mismos sitios y fechas).
8. Aprovechamiento de `boxplot.umd.min.js` para mostrar distribuciones por estado.
9. Etiquetado de versiones (releases) por cada actualización de datos.

---

# 14. Glosario

| Término | Definición |
|---|---|
| **APTA / NO APTA** | Clasificación sanitaria según el criterio recreativo (≤200 NMP/100 mL de enterococos) |
| **Arrastre** | Heredar hacia abajo el último valor no vacío, para interpretar celdas combinadas |
| **Coroplético** | Mapa que colorea regiones según el valor de una variable |
| **Enterococos fecales** | Bacterias indicadoras de contaminación fecal en agua |
| **Formato columnar** | Almacenamiento en arreglos paralelos indexados por registro |
| **Idempotente** | Que ejecutado varias veces produce siempre el mismo resultado |
| **Límite de detección** | Valor mínimo/máximo cuantificable por el laboratorio; se reporta como `<3`, `>24196` |
| **NMP/100 mL** | Número Más Probable por 100 mililitros: unidad de concentración bacteriana |
| **Operativo** | Campaña de muestreo de una temporada (Semana Santa, Verano o Invierno) |
| **Robot** | El flujo automatizado de GitHub Actions |
| **Sitio de muestreo** | Punto físico donde se toma la muestra; varios sitios forman una playa |
| **Sitio estático** | Web servida como archivos, sin servidor de aplicación ni base de datos |
| **`window.DEFAULT_DATA`** | Variable global con los datos de playas |
| **`window.IND_DATA`** | Variable global con los datos de indicadores |
| **`__name__ == "__main__"`** | Condición que hace que un script Python se ejecute solo al invocarlo directamente |

---

## Síntesis

El sistema convierte archivos Excel oficiales en tableros web interactivos mediante tres etapas
desacopladas: **conversores en Python** que limpian y normalizan los datos aplicando reglas
explícitas y verificables; **análisis en el navegador** que aplica el criterio sanitario y la regla
conservadora de agregación por playa; y un **flujo automatizado** que enlaza ambas etapas de modo
que publicar información nueva se reduce a subir un archivo. La validación por reconstrucción del
estado publicado garantiza que la automatización preserva exactamente la información original.
