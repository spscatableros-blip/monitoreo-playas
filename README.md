# Tablero CONAGUA · Calidad del agua en sitios costeros

Tablero web para consultar la calidad del agua en playas e indicadores costeros de México.
Publicado en: **https://spscatableros-blip.github.io/monitoreo-playas/**

Es un sitio **estático**: solo HTML, CSS y JavaScript. No requiere servidor, base de datos ni
instalación — basta abrir la página en cualquier navegador.

---

## Contenido

El tablero tiene dos secciones independientes, accesibles por pestañas:

### 🏖 Playas

Monitoreo bacteriológico de playas (enterococos fecales) del programa de vigilancia
CONAGUA · COFEPRIS.

| | |
|---|---|
| **Periodo** | 2014–2026 (33 operativos) |
| **Cobertura** | 17 estados costeros · 87 destinos turísticos · 293 playas |
| **Criterio** | Una playa es **NO APTA** si excede 200 NMP/100 mL de enterococos |
| **Operativos** | Semana Santa, Verano e Invierno de cada año |

Secciones: informe automático de la situación (nacional o por estado), sitios en alerta,
resultados nacionales por año, playas NO APTA por año y temporada, cumplimiento por estado,
ranking de destinos con más incumplimientos, mapa de calor por estado/año/periodo, sitios con
incumplimiento recurrente, explorador por estado y mapa nacional.

### 🧪 Indicadores costeros

Indicadores fisicoquímicos y bacteriológicos de cuerpos de agua costeros.

| | |
|---|---|
| **Periodo** | 2012–2025 |
| **Cobertura** | 17 estados · 950 sitios · 23,447 muestras |
| **Indicadores** | Enterococos fecales, sólidos suspendidos totales (SST) y oxígeno disuelto (% y mg/L, en superficie, media y fondo) |

Secciones: informe de resultados, estados en alerta, tendencia por año, comparación por estado,
cumplimiento de enterococos, mapa de calor estado × año y categorización por mes.

---

## Cómo actualizar los datos de playas

El tablero permite actualizar los datos **sin tocar código**, desde la propia página:

1. Pulsa **📄 Plantilla de ejemplo** para descargar un `.xlsx` con el formato correcto.
2. Llena la plantilla con los datos nuevos (o descarga los actuales con **⬇ Datos .xlsx** y edítalos).
3. Pulsa **⬆ Subir datos (.xlsx)** y selecciona el archivo. Todas las gráficas se recalculan al instante.

> La actualización es temporal: vive solo en tu navegador durante esa sesión. Para dejarla
> publicada de forma permanente hay que regenerar `playas/data.js` y hacer commit (ver más abajo).

### Formato del archivo

**Cada hoja = un operativo**, nombrada `temporada + año`:
`semana.santa26`, `verano24`, `invierno2021` (el año admite 2 o 4 dígitos).

**Columnas** (primera fila de cada hoja; el orden no importa, se detectan por nombre):

| Estado | Destino | Playa | Sitio | Clasificación |
|---|---|---|---|---|
| Baja California | Tijuana | Tijuana | Tijuana I | Apta |
| | | | Tijuana II | No apta |
| Colima | Manzanillo | Miramar | Miramar I | Apta |

- **Estado / Destino / Playa** pueden dejarse en blanco en filas seguidas de la misma playa
  (heredan el último valor no vacío, igual que las celdas combinadas del archivo oficial).
- Solo **Sitio** es obligatorio en cada fila; las filas sin sitio se ignoran.
- **Clasificación** acepta cualquier texto que contenga `apta` o `no apta` (sin distinguir
  mayúsculas). Otro texto o celda vacía se interpreta como sin dato.

El propio tablero incluye esta guía en la sección desplegable *"Formato requerido para subir datos"*.

---

## Estructura del repositorio

```
├── index.html              Página raíz con las pestañas (carga las dos secciones en iframes)
├── playas/
│   ├── index.html          Tablero de playas
│   ├── data.js             Datos 2014–2026 (window.DEFAULT_DATA)
│   ├── mexico.js           Geometría de los estados para el mapa
│   ├── assets/mapa.png     Mapa editorial de referencia
│   └── lib/                Chart.js, datalabels, annotation, SheetJS (xlsx)
└── indicadores/
    ├── index.html          Tablero de indicadores
    ├── data.js             Datos 2012–2025 (window.IND_DATA)
    └── lib/                Chart.js, datalabels, boxplot
```

Las librerías están **incluidas en el repositorio** (no se cargan desde un CDN), así que el
tablero funciona incluso sin conexión a internet.

---

## Formato de los datos

### `playas/data.js`

```js
window.DEFAULT_DATA = {
  periods: [
    {
      key: "verano24",           // identificador de la hoja
      label: "Verano 2024",      // se muestra en las gráficas; debe incluir el año
      rows: [
        { est: "Colima", dst: "Manzanillo", playa: "Miramar",
          sitio: "Miramar I", apta: 1 }   // 1 = APTA · 0 = NO APTA · null = sin dato
      ]
    }
  ]
}
```

### `indicadores/data.js`

Formato columnar (arreglos paralelos, un índice por muestra):

```js
window.IND_DATA = {
  estados: ["BAJA CALIFORNIA", ...],   // catálogo
  est:  [0, 0, 1, ...],                // índice al catálogo de estados
  anio: [2012, 2013, ...],
  mes:  [1, 2, ...],                   // 1–12
  indicadores: ["ENTEROC_FEC", "SST", ...],
  v: { ENTEROC_FEC: [10.0, null, ...], SST: [...] },
  meta: { muestras, sitios, anio_min, anio_max, umbral_enteroc, fuente }
}
```

---

## Desarrollo local

Al usar rutas relativas basta con abrir `index.html` en el navegador. Si tu navegador bloquea
la carga de archivos locales, levanta un servidor:

```bash
python3 -m http.server 8000
# abrir http://localhost:8000
```

### Publicar cambios

El sitio se publica con **GitHub Pages** desde la rama `main`. Cualquier cambio subido con
`git push` queda en línea en 1–2 minutos, sin pasos adicionales.

---

## Notas sobre los datos

- **2014–2026 (playas):** transcritos y verificados contra fuentes oficiales de SEMARNAT,
  tabla por tabla.
- **2022–2023 (playas):** provienen de una extracción previa que **no pudo verificarse** contra
  una fuente oficial. Úsense con reserva.
- **2013 (playas):** excluido del tablero. datos  disponibles provenían de un PDF incompleto y no era confiable.
- **Celdas vacías en los mapas de calor:** corresponden a operativos sin resultados reportados
  por la fuente (`SIN RESULTADOS`, `NO DISPONIBLE`, o sin muestreo — por ejemplo Acapulco en
  Invierno 2023 tras el huracán Otis). No se cuentan como aptas ni como no aptas, para no
  distorsionar los porcentajes de cumplimiento.
- **Nombres de playa entre años:** un mismo sitio puede aparecer con distinta grafía según el
  año (por ejemplo `Playa de Rosarito` en 2014–2021 y `Rosarito` en 2024–2026). Es intencional:
  cada año conserva la nomenclatura de su fuente oficial.

---

## Fuentes

- CONAGUA · COFEPRIS — Programa de Playas Limpias (monitoreo bacteriológico de playas)
- CONAGUA — Base de datos de calidad del agua en cuerpos costeros
- SEMARNAT — Tablas oficiales de monitoreo por destino y periodo
