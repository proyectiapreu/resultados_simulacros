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

## Identidad visual de las evaluaciones

- Recopilación Examen 2026-II: morado.
- Simulacro I: amarillo.
- Simulacro II: azul.
- Simulacro III: rojo.
- Simulacro IV: verde.

Estos colores identifican las tarjetas de acceso, los reportes individuales y los puntos de las gráficas.

## Presentación de puntajes

- `puntaje_global` se lee directamente del Excel y se muestra sin decimales mediante truncamiento visual.
- Los puntajes de las áreas se leen directamente del Excel y se muestran con una cifra decimal.
- Los aciertos no se usan para calcular el puntaje global ni los puntajes por área.
- La página no clasifica los resultados mediante categorías cualitativas ni fija metas de puntaje.
- Cada prueba se trata como una aplicación independiente con dificultad diseñada para ser similar. Las gráficas permiten comparar desempeños entre aplicaciones, no constituyen una equiparación estadística entre formas.
- En las gráficas aparecen siempre las cinco evaluaciones en el eje horizontal. Solo se dibujan puntos para las evaluaciones que el estudiante haya presentado y cuyos resultados estén disponibles.

## Hojas obligatorias del Excel

### `resultados`

Una fila por estudiante. Columnas principales:

`documento`, `nombre`, `evaluacion`, `orden`, `fecha`, `puntaje_global`, `aciertos_totales`, `total_preguntas`, `puntaje_analisis_textual`, `puntaje_matematicas`, `puntaje_ciencias_naturales`, `puntaje_ciencias_sociales`, `puntaje_analisis_imagen`, `aciertos_analisis_textual`, `aciertos_matematicas`, `aciertos_ciencias_naturales`, `aciertos_ciencias_sociales`, `aciertos_analisis_imagen`, `aciertos_fisica`, `aciertos_quimica`, `aciertos_biologia`, `total_fisica`, `total_quimica`, `total_biologia`.

`nombre` aparece dentro del reporte del estudiante. Los campos `total_fisica`, `total_quimica` y `total_biologia` permiten mostrar resultados como `6/11`.

Las columnas antiguas `area_prioritaria` y `fortaleza_principal` pueden permanecer en archivos ya existentes, pero la interfaz ya no las utiliza.

### `areas`

`documento`, `nombre`, `evaluacion`, `orden`, `area`, `puntaje_area`, `aciertos`, `total_preguntas`.

### `temas`

`documento`, `nombre`, `evaluacion`, `orden`, `area`, `subarea`, `tema`, `aciertos`, `errores`, `total_preguntas`, `porcentaje_acierto`.

La sección de progresión temática permite elegir un área y un tema. Muestra todas las cinco evaluaciones en el eje horizontal y coloca puntos únicamente en las aplicaciones donde ese tema tenga información para el estudiante. Debajo de la gráfica se muestran también los aciertos sobre el total de preguntas.

Para que un mismo tema se conecte correctamente entre pruebas, `area`, `subarea` y `tema` deben escribirse de manera consistente en todos los archivos.

### `preguntas`

`documento`, `nombre`, `evaluacion`, `orden`, `pregunta`, `area`, `subarea`, `tipo`, `tema`, `respuesta_estudiante`, `respuesta_correcta`, `resultado`.

El detalle por pregunta se presenta dentro de una ventana desplazable y muestra siempre la respuesta correcta, incluso cuando el estudiante respondió correctamente.

### `metadatos`

Incluye información descriptiva de la evaluación y la escala. No se utiliza para recalcular los puntajes.

## Recopilación Examen 2026-II

`data0/resultados.xlsx` integra los resultados de la recopilación y los nombres tomados de la hoja de respuestas original utilizada para esa aplicación.

## Perfil de muestra

El documento `1234` permite revisar la experiencia completa de la página. Tiene información disponible en la recopilación y en los cuatro simulacros, incluida la comparación global, la progresión por área, la progresión temática y el detalle por pregunta.

## Flujo de publicación

1. Se aplica la evaluación.
2. Se procesa la información y se genera el análisis Rasch fuera de la página.
3. Se exporta el Excel final con las cinco hojas anteriores.
4. Se reemplaza `resultados.xlsx` dentro de `data0`, `data1`, `data2`, `data3` o `data4`.
5. GitHub Pages publica la actualización.

## Dependencia

La página usa SheetJS 0.20.3 desde el CDN oficial únicamente para leer los archivos `.xlsx` en el navegador.
