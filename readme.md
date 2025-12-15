# Clustering K-Means en Excel (con técnica del codo)

Este repositorio contiene un ejemplo completo de **clustering K-Means implementado únicamente con herramientas de Office** (Excel + VBA), incluyendo:

- Un libro de Excel que ejecuta **K-Means** sobre una tabla de datos binarios (one-hot encoding).
- Un libro de Excel para aplicar la **técnica del codo** y ayudar a elegir el número óptimo de clusters.
- Datos de ejemplo para poder probar todo sin necesidad de preparar un dataset propio.

El objetivo es mostrar que se puede hacer **aprendizaje no supervisado** en Excel de forma transparente, viendo cada paso del algoritmo y sin depender de librerías externas.

---

## Contenido del repositorio

- `/Clusters/`  
  Libros de Excel con la implementación de **K-Means** en VBA:
  - Hoja con los datos (matriz de características, normalmente 0/1).
  - Rango con los centroides iniciales.
  - Columna donde se guarda el **cluster asignado** a cada fila.
  - Módulo VBA con el algoritmo de K-Means.

- `/Codo/`  
  Libros de Excel para aplicar la **técnica del codo**:
  - Ejecutan K-Means para distintos valores de K.
  - Calculan la **suma de cuadrados intra-cluster (SSE)**.
  - Generan un gráfico para visualizar el “codo” y ayudar a elegir el número de clusters.

- `/data/`  
  Datos de ejemplo utilizados en los libros de Excel. Están pensados para poder ejecutar el clustering directamente sin tener que preparar datos adicionales.

---

## Requisitos

- **Microsoft Excel de escritorio** (Windows).
- Macros **habilitadas** (los libros usan VBA).
- Permisos para ejecutar código VBA.

---

## Cómo usar el ejemplo de K-Means

1. Abre el libro de Excel correspondiente en la carpeta `/Clusters`.
2. Habilita macros cuando Excel te lo pida.
3. Revisa las hojas:
   - Rango de datos, por ejemplo algo tipo `B2:L181` (filas = observaciones, columnas = variables).
   - Rango de centroides iniciales, por ejemplo `B184:L188` (una fila por cluster).
   - Columna de asignación de cluster, por ejemplo `M2:M181`, donde se escriben valores como `C1`, `C2`, etc.
4. Ejecuta la macro desde:
   - **Desarrollador → Macros → `KMeansClustering` → Ejecutar**, o
   - El botón que haya en la hoja (si está configurado).
5. Al finalizar, verás:
   - Cada fila clasificada en un cluster (`C1`, `C2`, …).
   - Los centroides actualizados en la tabla de centros.
   - Un mensaje con el número de iteraciones realizadas.

---

## Detalles del algoritmo (VBA)

En el módulo VBA se implementa el algoritmo K-Means clásico:

1. **Inicialización**
   - Se leen:
     - El rango de datos.
     - El rango de centros (centroides).
     - El rango donde se escribe el cluster asignado.
   - El número de clusters se deduce a partir del **número de filas del rango de centros**.

2. **Asignación de clusters**
   - Para cada fila de la tabla:
     - Se calcula la **distancia euclídea** a cada centro.
     - Se asigna la fila al centro más cercano.
     - Se guarda el resultado como `C1`, `C2`, … en la columna de clusters.

3. **Actualización de centroides**
   - Para cada cluster:
     - Se hace la **media** de las filas asignadas a ese cluster, columna a columna.
     - Ese vector de medias se convierte en el nuevo centro.

4. **Criterio de parada**
   - Se repiten los pasos de asignación y actualización hasta que:
     - Las asignaciones dejan de cambiar, o
     - Se alcanza un número máximo de iteraciones (por ejemplo, 100).

Este enfoque permite ver de forma muy clara cómo funciona K-Means “por dentro”, sin caja negra.

---

## Técnica del codo en Excel

En los libros de la carpeta `/Codo` se aplica la **técnica del codo**:

1. Se ejecuta K-Means para distintos valores de K (por ejemplo, de 1 a 10).
2. Para cada K se calcula la **SSE (Sum of Squared Errors)**:
   - La suma de las distancias cuadráticas de cada punto a su centro.
3. Se dibuja un **gráfico SSE vs K**.
4. El “codo” del gráfico (donde la mejora empieza a ser marginal) indica un **buen valor de K**.

De esta forma puedes:
- Ver cómo cambia la compactación de los clusters según K.
- Justificar de forma visual el número de clusters elegido, incluso trabajando solo con Excel.

---

## Adaptar el ejemplo a tus propios datos

1. Copia uno de los libros de `/Clusters` y renómbralo.
2. Reemplaza la tabla de datos por tus propias variables (por ejemplo, one-hot encoding de motivos, productos, segmentos, etc.).
3. Ajusta:
   - El rango de datos en el código VBA (si cambian filas/columnas).
   - El rango de centros iniciales (una fila por cluster).
4. Si quieres cambiar el número de clusters:
   - Añade o quita filas en la tabla de centros.
   - Asegúrate de que el rango de centros del código cubre todas esas filas.

---

## Licencia

Este proyecto se publica con fines educativos y demostrativos.  
Puedes adaptarlo y reutilizarlo para tus propios experimentos de clustering en Excel.

---

