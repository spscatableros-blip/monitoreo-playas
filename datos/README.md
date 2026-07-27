# Carpeta de entrada de datos (playas)

Sube aquí tu **Excel (`.xlsx`)** con los datos de playas. El robot **detecta automáticamente**
el formato y regenera `playas/data.js`.

## Cómo actualizar el dashboard (automático)

1. Entra a esta carpeta en GitHub y pulsa **Add file → Upload files**.
2. Arrastra tu archivo `.xlsx` y confirma el commit.
3. Eso es todo. Un robot de GitHub Actions regenera `playas/data.js`, hace commit,
   y GitHub Pages publica el dashboard actualizado en 1–2 minutos.

No necesitas instalar nada ni tocar código.

## Se aceptan dos formatos

### A) Reporte SEMARNAT (recomendado) — una hoja por año

El Excel maestro tipo *"Monitoreo Playas SEMARNAT"*: **una hoja por año** (2014, 2024, …),
con columnas `Estado`, `Destino turístico`, `Playa`, `Sitio de muestreo` y un bloque de
3 columnas por temporada (`Fecha`, `NMP/100 mL`, `Clasificación`), cuyo encabezado dice
*"Monitoreo prevacacional {Semana Santa|Verano|Invierno} {año}"*.

- Es un **volcado completo**: **reemplaza** todos los datos de playas.
- El año **2013** se **omite** por defecto (el reporte lo marca como preliminar).
- `Estado / Destino / Playa` pueden venir combinadas (se arrastran).

### B) Por temporada — una hoja por operativo

- **Cada hoja = una temporada**, nombrada `semana.santa26`, `verano24`, `invierno2021`, etc.
- **Columnas**: `Estado`, `Destino`, `Playa`, `Sitio`, `Clasificación`.
- Es **incremental**: agrega/actualiza solo las temporadas presentes; el resto no se toca.

En ambos: `Clasificación` acepta texto que contenga `apta` → apta, `no apta` → no apta;
vacío u otro → sin dato.
