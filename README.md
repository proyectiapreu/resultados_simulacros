# Proyectia · Resultados UNAL y simulacros

Sitio estático listo para GitHub Pages. Consulta cinco archivos de Excel independientes y muestra los puntajes que ya fueron calculados previamente con el procedimiento de calificación correspondiente. El navegador no reconstruye ni recalcula puntajes Rasch.

## Estructura

```text
index.html
style.css
app.js
assets/
  logo-proyectia-morado.png
  logo-proyectia-claro.png
  social/
data0/resultados.xlsx   ← Recopilación Examen 2026-II
data1/resultados.xlsx   ← Simulacro I
data2/resultados.xlsx   ← Simulacro II
data3/resultados.xlsx   ← Simulacro III
data4/resultados.xlsx   ← Simulacro IV
```

Para publicar un nuevo simulacro se reemplaza únicamente el archivo `resultados.xlsx` de la carpeta correspondiente.

## Presentación de puntajes

- `puntaje_global` se lee directamente del Excel y se muestra sin decimales mediante truncamiento visual.
- Los puntajes de las áreas se leen directamente del Excel y se muestran con una cifra decimal.
- Los aciertos no se usan para calcular el puntaje global ni los puntajes por área.
- La progresión compara los puntajes ya guardados en cada archivo.

## Hojas obligatorias del Excel

### `resultados`

Una fila por estudiante. Columnas:

`documento`, `nombre`, `evaluacion`, `orden`, `fecha`, `puntaje_global`, `aciertos_totales`, `total_preguntas`, `puntaje_analisis_textual`, `puntaje_matematicas`, `puntaje_ciencias_naturales`, `puntaje_ciencias_sociales`, `puntaje_analisis_imagen`, `aciertos_analisis_textual`, `aciertos_matematicas`, `aciertos_ciencias_naturales`, `aciertos_ciencias_sociales`, `aciertos_analisis_imagen`, `aciertos_fisica`, `aciertos_quimica`, `aciertos_biologia`, `area_prioritaria`, `fortaleza_principal`, `total_fisica`, `total_quimica`, `total_biologia`.

`nombre` aparece dentro del reporte del estudiante. Los campos `total_fisica`, `total_quimica` y `total_biologia` permiten mostrar resultados como `6/11`.

### `areas`

`documento`, `nombre`, `evaluacion`, `orden`, `area`, `puntaje_area`, `aciertos`, `total_preguntas`.

### `temas`

`documento`, `nombre`, `evaluacion`, `orden`, `area`, `subarea`, `tema`, `aciertos`, `errores`, `total_preguntas`, `porcentaje_acierto`.

Los nombres de los temas deben mantenerse consistentes entre evaluaciones si se quiere comparar su progresión.

### `preguntas`

`documento`, `nombre`, `evaluacion`, `orden`, `pregunta`, `area`, `subarea`, `tipo`, `tema`, `respuesta_estudiante`, `respuesta_correcta`, `resultado`.

El detalle por pregunta se presenta dentro de una ventana desplazable para evitar que el reporte se extienda excesivamente.

### `metadatos`

Incluye información descriptiva de la evaluación y la escala. No se utiliza para recalcular los puntajes.

## Recopilación Examen 2026-II

`data0/resultados.xlsx` integra los resultados de la recopilación y los nombres tomados de la hoja de respuestas original utilizada para esa aplicación.

## Perfil de muestra

El documento `1234` permite revisar la experiencia completa de la página. Tiene información disponible en la recopilación y en los cuatro simulacros, incluida la progresión general, la progresión por área, los temas y el detalle por pregunta.

## Flujo de publicación

1. Se aplica la evaluación.
2. Se procesa la información y se genera el análisis Rasch fuera de la página.
3. Se exporta el Excel final con las cinco hojas anteriores.
4. Se reemplaza `resultados.xlsx` dentro de `data0`, `data1`, `data2`, `data3` o `data4`.
5. GitHub Pages publica la actualización.

## Dependencia

La página usa SheetJS 0.20.3 desde el CDN oficial únicamente para leer los archivos `.xlsx` en el navegador.


## Ajustes de interfaz v4

- La progresión temática fue retirada.
- La progresión se concentra en puntaje global y puntajes por área.
- Los resultados por área muestran una sola vez el número de respuestas correctas.
- La página incluye una advertencia institucional al final, antes de las redes sociales.
- El botón flotante de progresión permanece disponible en el menú y en los reportes individuales.
- La interfaz está adaptada para escritorio y dispositivos móviles.
