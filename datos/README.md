# Carpeta de entrada de datos

Sube aquí tus archivos **Excel (`.xlsx`)** con los operativos de playas.

## Cómo actualizar el dashboard (automático)

1. Entra a esta carpeta en GitHub y pulsa **Add file → Upload files**.
2. Arrastra tu archivo `.xlsx` (o varios) y confirma el commit.
3. Eso es todo. Un robot de GitHub Actions:
   - lee tus Excel,
   - regenera `playas/data.js`,
   - hace commit del resultado,
   - y GitHub Pages publica el dashboard actualizado en 1–2 minutos.

No necesitas instalar nada ni tocar código.

## Formato del Excel

- **Cada hoja = una temporada**, nombrada `semana.santa26`, `verano24`, `invierno2021`, etc.
  (empieza por `semana santa` / `verano` / `invierno` seguido del año en 2 o 4 dígitos).
- **Columnas** (primera fila, el orden no importa): `Estado`, `Destino`, `Playa`, `Sitio`, `Clasificación`.
- `Estado / Destino / Playa` pueden dejarse en blanco en filas seguidas de la misma playa
  (heredan el último valor). Solo `Sitio` es obligatorio en cada fila.
- `Clasificación`: texto que contenga `apta` → apta, `no apta` → no apta; vacío u otro → sin dato.

> La combinación es **incremental**: las temporadas nuevas se agregan y las que ya existían
> (misma hoja) se actualizan. El resto de temporadas no se toca. Puedes dejar tus Excel aquí como
> historial; el robot los procesa todos cada vez.
