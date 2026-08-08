# Proyectia · Resultados UNAL y simulacros

Sitio estático listo para GitHub Pages. La página consulta cinco archivos de Excel independientes y presenta los puntajes ya calculados. **El navegador no calcula ni transforma puntajes Rasch.**

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

Para publicar un nuevo simulacro solo hay que reemplazar el archivo `resultados.xlsx` de su carpeta. No se modifica el HTML, el CSS ni el JavaScript.

## Escala que debe venir calculada desde el análisis

La Universidad Nacional de Colombia califica la prueba mediante Teoría de Respuesta al Ítem y modelo de Rasch. Los puntajes por componente se reportan en una escala centrada en 10 y el puntaje total en una escala centrada en 500. El Excel que se suba debe contener esos puntajes ya calculados por el procedimiento utilizado para cada simulacro.

La página **solo lee**:

- `puntaje_global`;
- `puntaje_analisis_textual`;
- `puntaje_matematicas`;
- `puntaje_ciencias_naturales`;
- `puntaje_ciencias_sociales`;
- `puntaje_analisis_imagen`.

No usa los aciertos para reconstruir los puntajes.

## Hojas obligatorias del Excel

### 1. `resultados`

Una fila por estudiante. Columnas:

`documento`, `nombre`, `evaluacion`, `orden`, `fecha`, `puntaje_global`, `aciertos_totales`, `total_preguntas`, `puntaje_analisis_textual`, `puntaje_matematicas`, `puntaje_ciencias_naturales`, `puntaje_ciencias_sociales`, `puntaje_analisis_imagen`, `aciertos_analisis_textual`, `aciertos_matematicas`, `aciertos_ciencias_naturales`, `aciertos_ciencias_sociales`, `aciertos_analisis_imagen`, `aciertos_fisica`, `aciertos_quimica`, `aciertos_biologia`, `area_prioritaria`, `fortaleza_principal`.

`nombre` se muestra dentro del reporte cuando está disponible.

### 2. `areas`

Una fila por estudiante y área. Columnas:

`documento`, `nombre`, `evaluacion`, `orden`, `area`, `puntaje_area`, `aciertos`, `total_preguntas`.

### 3. `temas`

Una fila por estudiante y tema. Columnas:

`documento`, `nombre`, `evaluacion`, `orden`, `area`, `subarea`, `tema`, `aciertos`, `errores`, `total_preguntas`, `porcentaje_acierto`.

Esta hoja alimenta las recomendaciones temáticas y la comparación de temas entre evaluaciones. Los nombres de los temas deben mantenerse iguales entre simulacros cuando se quiera comparar su progreso.

### 4. `preguntas`

Una fila por estudiante y pregunta. Columnas:

`documento`, `nombre`, `evaluacion`, `orden`, `pregunta`, `area`, `subarea`, `tipo`, `tema`, `respuesta_estudiante`, `respuesta_correcta`, `resultado`.

### 5. `metadatos`

Incluye información de la evaluación y de la escala. La página no usa esta hoja para recalcular puntajes.

## Recopilación Examen 2026-II

`data0/resultados.xlsx` contiene los puntajes actualmente publicados en el repositorio de la recopilación 2026-II, junto con el detalle de respuestas y temas reconstruido a partir de los archivos públicos de esa recopilación.

El repositorio público de la recopilación no contiene nombres de estudiantes. Por esa razón la columna `nombre` de `data0` queda vacía. Para mostrar el nombre también en la recopilación, hay que completar esa columna con la base original de inscripción y repetir el mismo nombre en las hojas `areas`, `temas` y `preguntas`.

## Publicación progresiva

1. Se aplica el simulacro.
2. Se ejecuta el análisis con el modelo de Rasch fuera de la página.
3. Se genera el Excel final con las cinco hojas anteriores.
4. Se reemplaza, por ejemplo, `data2/resultados.xlsx` para publicar el Simulacro II.
5. GitHub Pages actualiza la página. El botón del simulacro queda disponible automáticamente para quienes tengan una fila en ese Excel.

## Dependencia

La página usa SheetJS 0.20.3 desde el CDN oficial para leer archivos `.xlsx` en el navegador. No se utiliza para calcular puntajes.
