# Cómo funciona el sistema — documentación técnica completa

Proyecto: **Tablero CONAGUA · Calidad del agua en sitios costeros**
(Monitoreo de Playas e Indicadores Costeros)

Este documento explica, a detalle y de principio a fin: **cómo se generan y transforman los
datos**, **cómo se analizan**, **cómo se genera el tablero**, qué hace **cada script**, cómo se
**llaman entre sí**, qué **librerías** se usan y cómo funciona la **automatización**.

---

## Índice

1. Arquitectura general
2. Estructura completa del repositorio
3. El formato de los datos (`data.js`)
4. Los conversores (`scripts/`) — función por función
5. Cómo se relacionan y se llaman las funciones
6. Cómo se analizan los datos (en el navegador)
7. Las librerías y el orden de carga
8. Cómo se genera el tablero (render)
9. La automatización (el robot)
10. Cómo actualizar los datos en la práctica
11. Glosario

---

## 1. Arquitectura general

Es un **sitio 100 % estático**: solo HTML, CSS y JavaScript. **No hay servidor ni base de datos.**
Todo el cálculo ocurre en el navegador. GitHub Pages solo sirve archivos.

El sistema tiene tres grandes momentos:

```
   PREPARAR                    ANALIZAR                    PUBLICAR
 (scripts Python)         (JavaScript, navegador)      (robot + Pages)
        │                          │                          │
  Excel → data.js          data.js → gráficas         commit → GitHub Pages
```

Idea central por cada tablero: **`data.js` = los datos**, **`index.html` = cómo se ven**. Separados.
Hay dos tableros (**playas** e **indicadores**) unidos por el `index.html` de la raíz con pestañas.

---

## 2. Estructura completa del repositorio

```
monitoreo-playas/
│
├── index.html                  PÁGINA RAÍZ: pestañas + iframes (une los 2 tableros)
├── README.md                   Documentación general
├── .gitignore                  Archivos que Git debe ignorar
│
├── playas/                     ── TABLERO 1: Playas ──
│   ├── index.html              Lógica + render del tablero de playas
│   ├── data.js                 DATOS de playas (window.DEFAULT_DATA)  ← se regenera
│   ├── mexico.js               Geometría de los estados (para el mapa)
│   ├── assets/mapa.png         Imagen de referencia
│   └── lib/                     Librerías locales:
│       ├── chart.umd.min.js       Chart.js (motor de gráficas)
│       ├── annotation.min.js      plugin de líneas/zonas
│       ├── datalabels.min.js      plugin de etiquetas numéricas
│       └── xlsx.full.min.js       SheetJS (leer Excel en el navegador)
│
├── indicadores/                ── TABLERO 2: Indicadores costeros ──
│   ├── index.html              Lógica + render del tablero de indicadores
│   ├── data.js                 DATOS de indicadores (window.IND_DATA)  ← se regenera
│   └── lib/
│       ├── chart.umd.min.js       Chart.js
│       ├── datalabels.min.js      plugin de etiquetas
│       └── boxplot.umd.min.js     (presente, no referenciado actualmente)
│
├── datos/                      BUZÓN: Excel de playas (subir aquí dispara la actualización)
│   └── README.md
├── datos-indicadores/          BUZÓN: Excel de indicadores
│   └── README.md
│
├── scripts/                    ── CONVERSORES (Excel → data.js) ──
│   ├── xlsx_a_datajs.py         Playas · formato por temporada (una hoja = un operativo)
│   ├── semarnat_a_datajs.py     Playas · reporte SEMARNAT (una hoja por año)
│   └── xlsx_a_inddatajs.py      Indicadores · base CONAGUA
│
├── .github/workflows/
│   └── actualizar-datos.yml     EL ROBOT (GitHub Actions): automatiza todo
│
└── docs/
    └── COMO_FUNCIONA.md         Este documento
```

---

## 3. El formato de los datos (`data.js`)

### 3.1 Playas — `playas/data.js` → `window.DEFAULT_DATA`

Lista de **periodos** (temporadas); cada fila es un sitio de muestreo con su clasificación:

```js
window.DEFAULT_DATA = {
  periods: [
    {
      key:   "verano24",              // identificador (temporada + año, 2 dígitos)
      label: "Verano 2024",           // nombre visible
      rows: [
        { est:"Baja California", dst:"Rosarito", playa:"Rosarito", sitio:"Rosarito I",  apta:1 },
        { est:"Baja California", dst:"Rosarito", playa:"Rosarito", sitio:"Rosarito II", apta:0 }
      ]
    }
  ]
}
```

| Campo   | Significado                  | Valores |
|---------|------------------------------|---------|
| `est`   | Estado                       | texto |
| `dst`   | Destino turístico            | texto |
| `playa` | Playa                        | texto |
| `sitio` | Sitio de muestreo            | texto |
| `apta`  | Clasificación bacteriológica | `1`=apta, `0`=no apta, `null`=sin dato |

### 3.2 Indicadores — `indicadores/data.js` → `window.IND_DATA`

Formato **columnar**: arreglos paralelos donde el índice `i` es una muestra.

```js
window.IND_DATA = {
  estados: ["BAJA CALIFORNIA", ...],       // catálogo de 17 estados costeros
  est:  [0, 0, 1, ...],                    // índice al catálogo, por muestra
  anio: [2012, 2012, ...],                 // año, por muestra
  mes:  [10, 10, ...],                     // mes 1–12, por muestra
  indicadores: ["ENTEROC_FEC","SST",...],  // nombres de los 8 indicadores
  v: { ENTEROC_FEC:[3.0,null,...], SST:[28,60,...], ... },   // valores por indicador
  meta: { muestras, sitios, anio_min, anio_max, umbral_enteroc:200, fuente }
}
```

La muestra `i` pertenece al estado `estados[est[i]]`, año `anio[i]`, mes `mes[i]`, con enterococos
`v.ENTEROC_FEC[i]`. Es compacto porque son ~23 000 muestras.

---

## 4. Los conversores (`scripts/`) — función por función

Su trabajo es **transformar el Excel en el `data.js`**. Todos comparten el esqueleto:
`main()` (lee opciones) → funciones de parseo → escribe el `data.js`.

Requisito común: `pip install openpyxl` (única librería externa; el resto es Python estándar).

### 4.1 `xlsx_a_datajs.py` — Playas, formato "por temporada"

Para Excel donde **cada hoja es una temporada** (`verano24`, `semana.santa26`…).

| Función             | Rol |
|---------------------|-----|
| `norm(s)`           | normaliza texto (minúsculas, sin acentos, sin espacios) para comparar encabezados |
| `g(v)`              | lee una celda como texto recortado (`""` si vacía) |
| `pretty_period(n)`  | nombre de hoja → etiqueta: `verano24` → `Verano 2024` (arma año de 4 dígitos) |
| `clean_sheet(ws)`   | una hoja → lista de filas `{est,dst,playa,sitio,apta}` |
| `cargar_data_js()`  | lee el `data.js` actual (para **mezclar**) |
| `escribir_data_js()`| escribe el `data.js` final |
| `main()`            | orquesta: lee cada hoja, mezcla y escribe |

**Reglas de `clean_sheet` (la transformación):**
1. Busca la fila de encabezado en las primeras 5 filas y localiza columnas **por nombre**
   (`Estado`, `Destino`, `Playa`, `Sitio`, `Clasificación`), no por posición.
2. `Estado/Destino/Playa` se **arrastran** hacia abajo (celda vacía hereda la de arriba → maneja
   celdas combinadas).
3. `Sitio` es obligatorio; filas sin sitio se ignoran.
4. `Clasificación` → `apta`: contiene `"no apta"`→`0`; contiene `"apta"`→`1`; vacío/otro→`null`.

**Comportamiento:** **incremental** — mezcla con el `data.js` existente por `key` de hoja
(temporada nueva se agrega; temporada repetida se reemplaza; las demás no se tocan). Crea un
respaldo `data.js.bak.<fecha>` salvo `--sin-respaldo`.

### 4.2 `semarnat_a_datajs.py` — Playas, "reporte SEMARNAT"

Para el Excel maestro donde **cada hoja es un AÑO** (`2014`, `2024`…), en formato ancho: columnas de
identificación y luego un **bloque de 3 columnas por temporada** (`Fecha`, `NMP/100 mL`,
`Clasificación`), con encabezado *"Monitoreo prevacacional {temporada} {año}"*.

| Función               | Rol |
|-----------------------|-----|
| `_norm(s)`            | normaliza texto |
| `_g(row,i)`           | lee una celda de forma segura |
| `parse_hoja_anio(ws)` | parsea una hoja-año → `{key: [filas]}` de cada temporada |
| `construir(...)`      | recorre todas las hojas-año, junta y ordena los periodos |
| `main()`              | orquesta y escribe |

**Cómo parsea `parse_hoja_anio` (4 pasos):**
1. **Encuentra la fila de encabezado**: la primera (de las 10 iniciales) que contiene `"estado"` y
   una columna que empieza con `"sitio"` (así ignora el título y filas vacías de arriba).
2. **Ubica las columnas** Estado/Destino/Playa/Sitio por nombre.
3. **Detecta los bloques de temporada**: busca en el encabezado el patrón
   `(Semana Santa|Verano|Invierno) + año`; cada coincidencia marca el inicio de un bloque, y su
   columna de `Clasificación` es la 3ª del bloque.
4. **Lee las filas** con **arrastre** de Estado/Destino/Playa; por cada bloque emite una fila
   `{est,dst,playa,sitio,apta}` (misma regla apta/no apta/null). Una fila del Excel genera una fila
   por cada temporada del año.

**Comportamiento:** **reemplazo completo** (el archivo es el maestro con todo el histórico). Omite
`2013` por defecto (el reporte lo marca preliminar); `--incluir-2013` lo incluye. Ordena
cronológicamente (año, y dentro del año: Semana Santa < Verano < Invierno).

**Ejemplo de arrastre + bloques.** Esta fila del Excel 2024:

| Estado | Destino | Playa | Sitio | …Verano: Clasif | …Invierno: Clasif |
|---|---|---|---|---|---|
| Baja California | Rosarito | Rosarito | Rosarito I | APTA | APTA |
| *(vacío)* | *(vacío)* | *(vacío)* | Rosarito II | APTA | NO APTA |

produce **4 filas** en el `data.js`:
- `verano24`   → `{…, sitio:"Rosarito I",  apta:1}`
- `verano24`   → `{…, sitio:"Rosarito II", apta:1}`  ← est/dst/playa arrastrados
- `invierno24` → `{…, sitio:"Rosarito I",  apta:1}`
- `invierno24` → `{…, sitio:"Rosarito II", apta:0}`  ← "NO APTA"

### 4.3 `xlsx_a_inddatajs.py` — Indicadores

Para el Excel de la base CONAGUA (hoja `indicadores`, una fila por muestra).

| Función                 | Rol |
|-------------------------|-----|
| `_norm(s)`              | normaliza texto |
| `numerizar(x)`          | valor → número: `'<3'`→`3`, `'>24196'`→`24196`, vacío/texto→`null` |
| `elegir_hoja(wb)`       | escoge la hoja `indicadores` |
| `localizar_columnas()`  | ubica por nombre: `CLAVE SITIO`, `FECHA REALIZACIÓN`, `Año`, `ESTADO`, y los 8 indicadores |
| `cargar_meta_previa()`  | conserva `umbral_enteroc` y `fuente` del `data.js` anterior |
| `construir(...)`        | arma los arreglos columnares |
| `main()`                | orquesta y escribe |

**Reglas de transformación:**
1. `ESTADO` se normaliza; si **no** es uno de los 17 estados costeros del catálogo, la fila se
   **descarta** (filtra estados no costeros y vacíos).
2. El **mes** sale del mes de `FECHA REALIZACIÓN`.
3. Cada indicador se numeriza con `numerizar`.
4. `meta.sitios` = `CLAVE SITIO` distintos; `meta.muestras` = filas válidas.

**Comportamiento:** reemplazo completo (el Excel es un volcado de toda la base). Validado como
**byte-idéntico** al `data.js` original (23 447 muestras, 950 sitios).

---

## 5. Cómo se relacionan y se llaman las funciones

Punto clave: **los 3 scripts son independientes entre sí.** Ninguno importa funciones de otro
(aunque algunos nombres coincidan, como `construir` o `_norm`: son funciones distintas con el mismo
nombre en cada archivo, no compartidas).

Dentro de cada script, las funciones se llaman **en cadena**, empezando por `main()`:

```
semarnat_a_datajs.py:               xlsx_a_datajs.py:                 xlsx_a_inddatajs.py:
  main()                              main()                            main()
   └─ construir()                      ├─ clean_sheet()                  ├─ cargar_meta_previa()
       └─ parse_hoja_anio()            ├─ pretty_period()                └─ construir()
           └─ _norm(), _g()            ├─ cargar_data_js()                   ├─ elegir_hoja()
                                       └─ escribir_data_js()                 ├─ localizar_columnas()
                                                                             └─ numerizar()
```

**¿Quién llama al script desde afuera?** Nadie llama a sus *funciones* desde otro archivo. Lo único
externo es el **robot**, que ejecuta el archivo **como programa**:

```
Robot (workflow)  →  python3 scripts/semarnat_a_datajs.py archivo.xlsx
                         │
                         ▼
                 if __name__ == "__main__":  →  main()   ← puerta de entrada
```

El bloque `if __name__ == "__main__": main()` hace que `main()` se ejecute **solo** cuando corres el
archivo directamente (no si se importara desde otro `.py`). Es el estándar de Python para separar
"programa ejecutable" de "módulo importable".

---

## 6. Cómo se analizan los datos (en el navegador)

**El análisis NO lo hacen los scripts de Python.** Ocurre en el JavaScript de cada `index.html`,
dentro del navegador, sobre `window.DEFAULT_DATA` / `window.IND_DATA`.

### 6.1 La regla central: cuándo una playa es "APTA"

Hay dos **niveles**: por `sitio` o por `playa` (el usuario alterna).

- **Nivel sitio:** cada sitio cuenta individualmente.
- **Nivel playa:** los sitios se agrupan por `est|dst|playa`, y **una playa es APTA solo si TODOS
  sus sitios son aptos** (`v.every(x => x === 1)`). Basta un sitio no apto para que la playa cuente
  como no apta. Es la regla central del tablero.

```js
// counts(rows): resume filas → {aptas, noaptas, total, pct}
const g = {};
for (const r of rows) (g[r.est+'|'+r.dst+'|'+r.playa] ??= []).push(r.apta);  // agrupa por playa
for (const k in g) {
  const v = g[k].filter(valid);      // valid(a) = (a===0 || a===1); ignora los null
  if (!v.length) continue;
  t++;                               // playa contable
  if (v.every(x => x === 1)) a++;    // apta solo si TODOS aptos
}
// pct = a / t * 100  → % de cumplimiento
```

### 6.2 Agregación por año

Cada año tiene hasta 3 temporadas. Para gráficas "por año":
- `periodYear(p)` extrae el año de la etiqueta.
- `periodsOfYear(y)` junta las temporadas de ese año.
- El % anual **promedia** los % de sus temporadas (`avgNN` = promedio ignorando `null`).

### 6.3 Qué produce el análisis (playas)

- **KPIs** del último operativo (estados, destinos, % de cumplimiento).
- **% de cumplimiento por año** (línea) y **playas aptas/no aptas** (barras apiladas).
- **Ranking de destinos** con más incumplimientos.
- **Mapa de calor** por estado × año/temporada (usa `mexico.js`).
- **Informe automático** en texto (nacional o por estado).

### 6.4 Indicadores

`indicadores/index.html` hace lo análogo sobre `window.IND_DATA`: filtra las muestras por
estado/año/mes cruzando los arreglos por índice, compara el enterococo fecal contra
`meta.umbral_enteroc` (200 NMP/100 mL) para clasificar cumplimiento, y grafica los indicadores
fisicoquímicos (oxígeno disuelto, sólidos suspendidos, etc.).

---

## 7. Las librerías y el orden de carga

Cada tablero incluye sus scripts con **rutas relativas** a archivos **locales** en `lib/`.
**No hay CDN: cero dependencias externas** → funciona sin internet y es reproducible. El navegador
ejecuta los `<script>` de arriba abajo, por eso **el orden es obligatorio**.

`playas/index.html` (líneas 7–12) carga, en orden:

| # | Archivo | Librería | Rol |
|---|---------|----------|-----|
| 1 | `lib/chart.umd.min.js`   | **Chart.js**              | motor de todas las gráficas (debe ir primero) |
| 2 | `lib/annotation.min.js`  | chartjs-plugin-annotation | líneas/zonas de umbral |
| 3 | `lib/datalabels.min.js`  | chartjs-plugin-datalabels | números encima de barras/puntos |
| 4 | `lib/xlsx.full.min.js`   | **SheetJS (xlsx)**       | leer Excel en el navegador (subida temporal) |
| 5 | `data.js`                | *(datos)*                 | deja `window.DEFAULT_DATA` |
| 6 | `mexico.js`              | *(datos)*                 | geometría de estados para el mapa |
| 7 | `<script>` inline (l.277)| *(código propio)*         | la lógica del tablero, que usa todo lo anterior |

`indicadores/index.html` (líneas 7–9) carga solo: `chart.umd.min.js`, `datalabels.min.js`,
`data.js`, y su `<script>` inline (l.162). *(`boxplot.umd.min.js` existe en `lib/` pero el HTML
actual no lo referencia.)*

**Por qué importa el orden:** Chart.js crea el objeto global `Chart`; los plugins se **registran**
dentro de él (lo necesitan cargado); `data.js` deja los datos en memoria; y al final el código
inline los **usa** (`lee DEFAULT_DATA → calcula → new Chart(...)`). Invertir el orden rompe todo.

---

## 8. Cómo se genera el tablero (render)

### 8.1 La página raíz (`index.html`)

No carga librerías. Su JavaScript muestra/oculta dos **iframes** (uno por tablero) y los carga
**perezosamente** (cada tablero se carga la primera vez que se abre su pestaña):

```
index.html (raíz)
 ├─ iframe → playas/index.html       (se carga al abrir "🏖 Playas")
 └─ iframe → indicadores/index.html  (se carga al abrir "🧪 Indicadores")
```

### 8.2 Secuencia de render de un tablero

1. **Carga las librerías** (`lib/`) → `Chart` y sus plugins listos.
2. **Carga los datos** (`data.js`; y `mexico.js` en playas).
3. **Ejecuta el análisis** (sección 6): `counts`, `pctRows`, etc.
4. **Dibuja** cada gráfica con `new Chart(...)` y pinta el mapa.
5. **Interactividad:** al cambiar de nivel (sitio/playa), estado o destino, recalcula y vuelve a
   dibujar (`destroy` + recrear).

Nada de esto necesita servidor: todo el cálculo ocurre en el navegador a partir del `data.js`.

---

## 9. La automatización (el robot)

Definida en `.github/workflows/actualizar-datos.yml`. Se dispara cuando subes/cambias un `.xlsx`
en `datos/` o `datos-indicadores/` (o manualmente desde la pestaña *Actions*).

```
1. Subes un Excel  →  datos/ (playas)  o  datos-indicadores/ (indicadores) en GitHub
                          │
2. GitHub Actions detecta el cambio (el robot)
                          │
3. Job "playas": DETECTA el formato del Excel y usa el conversor adecuado
     • hojas por año (2014, 2024...)   →  semarnat_a_datajs.py   (reemplazo)
     • hojas por temporada (verano24)  →  xlsx_a_datajs.py       (mezcla)
   Job "indicadores": xlsx_a_inddatajs.py  (reemplazo)
                          │
4. Si el data.js cambió, el robot hace commit (github-actions[bot])
                          │
5. GitHub Pages republica  →  el tablero muestra los datos nuevos en 1–2 min
```

**Cómo detecta el formato de playas:** abre el Excel y mira los nombres de las hojas; si alguna es
un año de 4 dígitos → reporte SEMARNAT; si no → formato por temporada.

**Detalles:** los dos jobs están **encadenados** (`indicadores` corre después de `playas`) para no
chocar al hacer `git push`. El commit del robot lleva `[skip ci]` para no re-dispararse a sí mismo.
Requisito de configuración: *Settings → Actions → Workflow permissions → Read and write*.

---

## 10. Cómo actualizar los datos en la práctica

**Playas** (recomendado: reporte SEMARNAT):
1. En GitHub, entra a `datos/` → **Add file → Upload files**.
2. Sube tu `Monitoreo Playas SEMARNAT ….xlsx`.
3. Confirma el commit. El robot regenera `playas/data.js` y publica.

**Indicadores:**
1. Entra a `datos-indicadores/` → **Add file → Upload files**.
2. Sube el Excel de la base (hoja `indicadores`).
3. Confirma. El robot regenera `indicadores/data.js` y publica.

**En tu computadora (opcional):**
```bash
pip install openpyxl
python3 scripts/semarnat_a_datajs.py "Monitoreo Playas SEMARNAT 2013-2026.xlsx" --dry-run  # previsualizar
python3 scripts/semarnat_a_datajs.py "Monitoreo Playas SEMARNAT 2013-2026.xlsx"            # generar
git add playas/data.js && git commit -m "Actualiza datos de playas" && git push            # publicar
```

---

## 11. Glosario

| Término | Significado |
|---|---|
| **Sitio estático** | Web de solo HTML/JS/CSS, sin servidor ni base de datos |
| **`data.js`** | Archivo que contiene los datos como variable global de JavaScript |
| **`window.DEFAULT_DATA`** | Los datos de playas |
| **`window.IND_DATA`** | Los datos de indicadores (formato columnar) |
| **Temporada / operativo** | Semana Santa, Verano o Invierno de un año |
| **`apta`** | Clasificación bacteriológica: 1=apta, 0=no apta, null=sin dato |
| **Arrastre** | Heredar hacia abajo el último valor no vacío (celdas combinadas) |
| **Formato columnar** | Datos en arreglos paralelos, indexados por muestra |
| **El robot** | El flujo de GitHub Actions que automatiza la actualización |
| **GitHub Pages** | El servicio que publica el sitio estático |
| **`if __name__ == "__main__"`** | Hace que un script Python se ejecute solo al correrlo directamente |

---

## Resumen en una frase

> El **Excel** se convierte en **`data.js`** mediante los **scripts de Python** (que el **robot**
> ejecuta automáticamente al subir el archivo a `datos/`); luego **`index.html`** lee ese `data.js`,
> **analiza** los datos en el navegador (regla: una playa es apta solo si todos sus sitios lo son) y
> **dibuja** el tablero con Chart.js — todo servido de forma estática por GitHub Pages.
