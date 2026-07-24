# Carpeta de entrada — Indicadores costeros

Sube aquí el **Excel completo de la base de calidad de sitios costeros** de CONAGUA
(la hoja debe llamarse `indicadores`). A diferencia de playas, este archivo es un
**volcado completo** de todas las muestras, así que **reemplaza** por completo los datos
del tablero de indicadores.

## Cómo actualizar (automático)

1. Entra a esta carpeta en GitHub y pulsa **Add file → Upload files**.
2. Sube tu Excel (`.xlsx`). Si ya había uno, súbelo con el mismo nombre para reemplazarlo
   (si subes varios, el robot usa el más reciente por orden de nombre).
3. Confirma el commit. El robot de GitHub Actions:
   - regenera `indicadores/data.js`,
   - hace commit,
   - y GitHub Pages publica el tablero actualizado en 1–2 minutos.

No necesitas instalar nada ni tocar código.

## Formato esperado del Excel

Una hoja `indicadores` con encabezados en la primera fila. El robot usa estas columnas
(las localiza **por nombre**, el orden no importa):

| Columna              | Uso |
|----------------------|-----|
| `CLAVE SITIO`        | contar sitios distintos (`meta.sitios`) |
| `FECHA REALIZACIÓN`  | de aquí sale el mes (1–12) |
| `Año`                | año de la muestra |
| `ESTADO`             | estado (debe ser uno de los 17 costeros; el resto se descarta) |
| `ENTEROC_FEC`, `SST`, `OD_%_SUP`, `OD_mg/L_SUP`, `OD_%_MED`, `OD_mg/L_MED`, `OD_%_FON`, `OD_mg/L_FON` | los 8 indicadores |

Notas:

- Las filas cuyo `ESTADO` no sea un estado costero del catálogo (p. ej. `SAN LUIS POTOSÍ`)
  o esté vacío **se descartan** automáticamente.
- Valores como `<3`, `<1`, `>24196` se convierten al número (`3`, `1`, `24196`). Celda vacía
  o texto no numérico → sin dato.
