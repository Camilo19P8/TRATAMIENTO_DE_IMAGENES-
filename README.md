# Informe de resultados de tratamiento de imagen

Práctica de Semana 8 · **Percepción Computacional** · Universidad de Cundinamarca.

**Autores:** Cristian Camilo Alarcon Delgadillo, Brayan Arley Saenz Cortes y Jonathan David Llanceros Bello.

**Docente:** Mónica María Fonseca Vigoya.

**Fecha del informe:** 12 de octubre de 2026.

## Contenido

- `datos/originales/`: las cinco fotografías propias, sin modificar.
- `src/practica.py`: diagnóstico, métricas, tratamiento y figuras.
- `src/proceso_visual.py`: figuras de las etapas intermedias, parámetros, bordes y detalles del resultado.
- `notebooks/Practica_tratamiento_imagenes.ipynb`: cuaderno para VS Code o Google Colab.
- `resultados/`: CSV, JSON, imagen tratada, figuras y capturas reales de VS Code.
- `src/generar_informe.py`: genera el documento editable a partir de resultados y capturas reales. La carpeta local `informe/` está excluida de GitHub; el Word y el PDF se entregan por separado.

## Ejecutar en VS Code

Abrir esta carpeta con **Archivo > Abrir carpeta** y seleccionar un intérprete Python 3.12 o superior. En una terminal del proyecto:

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe src/practica.py
.venv\Scripts\python.exe src/proceso_visual.py
```

También funciona en Linux y macOS usando `.venv/bin/python`. No es necesario activar el entorno. El programa escribe los resultados en `resultados/`, conserva los originales y no necesita una interfaz gráfica de OpenCV.

En esta copia local ya está preparado el entorno `.venv`. En VS Code, **Ctrl+Mayús+B** ejecuta la tarea de la práctica y **F5** usa la configuración de ejecución incluida. En otra copia del repositorio debe crearse primero el entorno con los comandos anteriores.

## Ejecutar en Google Colab

[Abrir el cuaderno en Colab](https://colab.research.google.com/github/Camilo19P8/TRATAMIENTO_DE_IMAGENES-/blob/main/notebooks/Practica_tratamiento_imagenes.ipynb)

Ejecutar las celdas en orden. La primera clona este repositorio cuando se ejecuta en Colab e instala las dependencias; las siguientes ejecutan el análisis, el proceso visual y muestran tablas y figuras. Se verificaron las seis celdas de código mediante ejecución local sin errores. Las salidas incluidas son reales y locales; el cuaderno no contiene salidas de Colab inventadas. Para ejecutar un notebook en VS Code se requiere también `ipykernel` y la extensión Jupyter.

## Resultado y limitaciones

La imagen I3 tiene una varianza global del Laplaciano de **54.23**, frente a **548.53** en I1. La máscara de enfoque en luminancia con cantidad 0.8 y sigma 1.2 eleva la respuesta de I3 a **149.05**. Esto realza los bordes existentes; no demuestra recuperación de información perdida ni mejora de un detector u OCR.

La foto recibida con el nombre **baja iluminación** tiene un brillo medio de **187.41**, superior a **124.66** en la referencia. Se conserva y se informa esta discrepancia. Para acreditar rigurosamente la condición de poca luz se necesita una nueva captura controlada; no se oscureció artificialmente ninguna fotografía para hacerla pasar por una captura real.

El objeto común es la lata negra Monster de la fila superior. Las coordenadas manuales de sus regiones se conservan en `resultados/manifiesto.json`. Para el diagnóstico entre fotografías se normalizan a 128 × 256; para el antes/después de I3 se usa el mismo recorte nativo. Son mediciones exploratorias: los cambios de perspectiva y escala impiden una comparación geométrica exacta.

## Métricas

- Resolución: ancho × alto en píxeles; tres canales BGR de 8 bits.
- Brillo: media del gris OpenCV, rango 0 a 255.
- Contraste: desviación estándar del gris.
- Nitidez: varianza del Laplaciano (`ksize=1`) y energía media Sobel (`ksize=3`, Tenengrad).
- Tamaño: bytes del archivo; el informe expresa KiB = bytes / 1024.
- Recorte del tratamiento: proporción de valores Y fuera de 0–255 antes de limitar; no equivale a saturación del gris final.

No se calculan PSNR ni SSIM entre tomas distintas al no disponer de una referencia alineada de I3.

## Regenerar el documento editable

```powershell
.venv\Scripts\python.exe -m pip install -r requirements_informe.txt
.venv\Scripts\python.exe src/generar_informe.py
```

Exportar el DOCX a PDF con Word o LibreOffice. El informe usa tamaño carta, márgenes de 2.54 cm, Times New Roman de 12 puntos, doble espacio en el cuerpo, numeración superior y portada académica. Las tablas tienen líneas internas visibles por solicitud de presentación, adaptación explícita respecto a las tablas APA.

## Fuentes

- [OpenCV: filtrado y suavizado](https://docs.opencv.org/4.13.0/d4/d13/tutorial_py_filtering.html)
- [OpenCV: ecualización e histogramas](https://docs.opencv.org/4.13.0/d5/daf/tutorial_py_histogram_equalization.html)
- [APA: ejemplo de trabajo estudiantil](https://apastyle.apa.org/style-grammar-guidelines/paper-format/student-annotated.pdf)
