"""Genera el informe editable usando resultados medidos y evidencias reales."""
from pathlib import Path
import json
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

ROOT = Path(__file__).resolve().parents[1]
DATA = json.loads((ROOT/'resultados/resumen.json').read_text(encoding='utf-8'))
FIG = ROOT/'resultados/figuras'
DOC = Document()
sec = DOC.sections[0]
sec.page_width, sec.page_height = Inches(8.5), Inches(11)
sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Inches(1)
sec.header_distance = Inches(.5)
for name in ['Normal', 'Title', 'Heading 1', 'Heading 2']:
    style = DOC.styles[name]
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    style.font.color.rgb = RGBColor(0,0,0)
    for fonts in style.element.iter(qn('w:rFonts')):
        for attr in ['asciiTheme','hAnsiTheme','eastAsiaTheme','cstheme']:
            fonts.attrib.pop(qn('w:'+attr),None)
    style.paragraph_format.line_spacing = 2
    style.paragraph_format.space_after = Pt(0)
DOC.styles['Normal'].paragraph_format.first_line_indent = Inches(.5)
DOC.styles['Title'].font.bold = True
DOC.styles['Title'].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
DOC.styles['Heading 1'].font.bold = True
DOC.styles['Heading 1'].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
DOC.styles['Heading 1'].paragraph_format.space_after = Pt(6)
DOC.styles['Heading 1'].paragraph_format.space_before = Pt(0)
for style in DOC.styles:
    for border in list(style.element.iter(qn('w:pBdr'))):
        border.getparent().remove(border)
header = sec.header.paragraphs[0]
header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
field = OxmlElement('w:fldSimple'); field.set(qn('w:instr'), 'PAGE')
header._p.append(field)
DOC.core_properties.title = 'Informe de resultados de tratamiento de imagen'
DOC.core_properties.author = 'Cristian Camilo Alarcon Delgadillo; Brayan Arley Saenz Cortes; Jonathan David Llanceros Bello'

def hyperlink(p,url):
    link=OxmlElement('w:hyperlink');link.set(qn('r:id'),p.part.relate_to(url,RT.HYPERLINK,is_external=True))
    run=OxmlElement('w:r');props=OxmlElement('w:rPr')
    fonts=OxmlElement('w:rFonts');fonts.set(qn('w:ascii'),'Times New Roman');fonts.set(qn('w:hAnsi'),'Times New Roman');props.append(fonts)
    run.append(props);txt=OxmlElement('w:t');txt.text=url;run.append(txt);link.append(run);p._p.append(link)

def para(text, indent=True, single=False):
    p=DOC.add_paragraph()
    if text.startswith('https://'):hyperlink(p,text)
    else:p.add_run(text)
    if not indent:p.paragraph_format.first_line_indent=Inches(0)
    if single:p.paragraph_format.line_spacing=1.15
    return p

def heading(text):DOC.add_heading(text,1)
def page(text):
    p=DOC.add_heading(text,1)
    p.paragraph_format.page_break_before=True
def fmt(n):return f'{n:.2f}'.replace('.',',')

def figure(num,title,path,width=6.3,note='Elaboración propia a partir de las fotografías suministradas.'):
    p=para(f'Figura {num}',False,True);p.runs[0].bold=True;p.paragraph_format.keep_with_next=True
    p=para(title,False,True);p.runs[0].italic=True;p.paragraph_format.keep_with_next=True
    p=DOC.add_paragraph();p.paragraph_format.first_line_indent=Inches(0)
    p.paragraph_format.line_spacing=1;p.paragraph_format.keep_with_next=True
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(path),width=Inches(width))
    p=para('Nota. '+note,False,True)
    for r in p.runs:r.font.size=Pt(10)
    p.paragraph_format.space_after=Pt(8)

def table(num,title,headers,rows,widths):
    p=para(f'Tabla {num}',False,True);p.runs[0].bold=True;p.paragraph_format.keep_with_next=True
    p=para(title,False,True);p.runs[0].italic=True;p.paragraph_format.keep_with_next=True
    t=DOC.add_table(rows=1,cols=len(headers));t.autofit=False;t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for c,w in zip(t.columns,widths):c.width=Inches(w)
    for c,h in zip(t.rows[0].cells,headers):c.text=h
    for row in rows:
        for cell,text in zip(t.add_row().cells,row):cell.text=str(text)
    repeat=OxmlElement('w:tblHeader');t.rows[0]._tr.get_or_add_trPr().append(repeat)
    for i,row in enumerate(t.rows):
        no_split=OxmlElement('w:cantSplit');row._tr.get_or_add_trPr().append(no_split)
        for j,cell in enumerate(row.cells):
            cell.width=Inches(widths[j]);cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            props=cell._tc.get_or_add_tcPr()
            borders=OxmlElement('w:tcBorders')
            for edge in ['top','left','bottom','right']:
                e=OxmlElement('w:'+edge);e.set(qn('w:val'),'single');e.set(qn('w:sz'),'8');e.set(qn('w:color'),'D9D9D9');borders.append(e)
            props.append(borders)
            shade=OxmlElement('w:shd');shade.set(qn('w:fill'),'DBE7EF' if i==0 else ('F4F7F9' if i%2==0 else 'FFFFFF'));props.append(shade)
            margins=OxmlElement('w:tcMar')
            for edge in ['top','left','bottom','right']:
                e=OxmlElement('w:'+edge);e.set(qn('w:w'),'80');e.set(qn('w:type'),'dxa');margins.append(e)
            props.append(margins)
            for p in cell.paragraphs:
                p.paragraph_format.first_line_indent=Inches(0);p.paragraph_format.line_spacing=1.1
                p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:r.font.size=Pt(10);r.bold=(i==0)
    para('',False,True)

def code(text):
    p=para(text,False,True);p.paragraph_format.line_spacing=1
    for r in p.runs:r.font.name='Consolas';r.font.size=Pt(9)

# Portada académica siguiendo la referencia suministrada.
p=DOC.add_paragraph();p.paragraph_format.space_after=Pt(66)
DOC.add_paragraph('Informe de resultados de tratamiento de imagen',style='Title')
p=para('Informe de práctica · Semana 8',False);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
p=DOC.add_paragraph();p.paragraph_format.space_after=Pt(24)
for text in ['Cristian Camilo Alarcon Delgadillo','Brayan Arley Saenz Cortes','Jonathan David Llanceros Bello']:
    p=para(text,False);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
p=DOC.add_paragraph();p.paragraph_format.space_after=Pt(24)
for text in ['Universidad de Cundinamarca','Especialización en Inteligencia Artificial','Percepción Computacional','Mónica María Fonseca Vigoya','12 de octubre de 2026']:
    p=para(text,False);p.alignment=WD_ALIGN_PARAGRAPH.CENTER

page('Propósito y metodología')
para('El diagnóstico de cinco fotografías de una misma escena permitió identificar una pérdida marcada de definición en la imagen I3. Aplicamos una máscara de enfoque moderada en luminancia y observamos un aumento de la respuesta de los bordes conservados. El detalle fino de las etiquetas continuó degradado; por ello, el resultado se interpreta como un realce visual parcial y no como una restauración completa.')
para('La práctica estudia cómo el encuadre, la iluminación y la estabilidad de la cámara afectan una imagen utilizada en percepción computacional. La lata negra Monster situada en la fila superior es el objeto común de las cinco tomas. Las demás latas, el teclado, las figuras y el computador aportan contexto y distractores. Los cinco archivos proceden de la carpeta de imágenes facilitada para esta actividad y se conservan sin modificar.')
para('Las condiciones se describen a partir de los nombres de los archivos y de la evidencia visible. No se dispone de registros de exposición, distancia, iluminación en lux ni dispositivo de captura. En particular, la condición denominada baja iluminación no se confirma mediante el brillo observado. Esta discrepancia se documenta para distinguir el nombre del archivo de su diagnóstico efectivo.')
para('Ejecutamos Python y OpenCV en el equipo local y organizamos el código en un proyecto de VS Code. El cuaderno complementario está preparado para ejecutar la misma práctica en Google Colab. Medimos resolución, canales, brillo, contraste, nitidez, formato y tamaño; comparamos los resultados y seleccionamos un tratamiento relacionado con el problema de I3. Las figuras muestran originales, regiones del objeto, histogramas y resultados antes y después.')

page('Fotografías de referencia e iluminación')
figure(1,'I1 con encuadre frontal y definición de referencia',FIG/'01_captura.png',5.25,'El rectángulo amarillo identifica la lata negra usada como objeto común. La escena incluye doce latas y elementos secundarios.')
figure(2,'I2 con mayor brillo observado',FIG/'02_captura.png',5.25,'Archivo original denominado baja iluminación. La imagen presenta mayor brillo que I1; el nombre no permite establecer la exposición física de la captura.')

page('Fotografías con desenfoque y fondo complejo')
figure(3,'I3 con bordes suaves y pérdida de detalle',FIG/'03_captura.png',5.25,'Se observan letras y contornos difusos. La fotografía por sí sola no permite separar con certeza desenfoque óptico y movimiento.')
figure(4,'I4 con múltiples distractores y menor escala del objeto',FIG/'04_captura.png',5.25,'El computador, la alfombrilla y los accesorios aumentan el contenido ajeno al objeto. Un contraste global alto no elimina esta dificultad.')

page('Perspectiva y comparación del mismo objeto')
figure(5,'I5 desde un ángulo elevado',FIG/'05_captura.png',5.8,'Las tapas son más visibles y cambia la proyección de la lata. La perspectiva modifica bordes, proporciones y zonas observables.')
figure(6,'Regiones del objeto en las cinco condiciones',FIG/'06_mismo_objeto.png',6.3,'De izquierda a derecha, I1 a I5. Recortes manuales normalizados a 128 × 256 píxeles con INTER_AREA. Esta normalización es exploratoria y no alinea la perspectiva.')

page('Resultados del diagnóstico')
rows=DATA['diagnostico']
table(1,'Propiedades de los archivos originales',['Imagen','Resolución\npx','Canales','Formato','Tamaño\nKiB'],[[f'I{i+1}',f"{r['ancho']} × {r['alto']}",r['canales'],r['formato'],fmt(r['bytes']/1024)] for i,r in enumerate(rows)],[.65,1.75,.85,.95,1.8])
para('Nota. KiB = bytes / 1024. Todas las imágenes se decodificaron en tres canales BGR y profundidad de 8 bits. I3 tiene una fila menos que las otras tomas; esta diferencia no explica la pérdida de definición.',False,True)
table(2,'Descriptores globales de brillo contraste y nitidez',['Imagen','Brillo\nμ','Contraste\nσ','Laplaciano\nVar','Tenengrad'],[[f'I{i+1}',fmt(r['brillo']),fmt(r['contraste']),fmt(r['laplaciano']),fmt(r['tenengrad'])] for i,r in enumerate(rows)],[.65,1.15,1.3,1.45,1.45])
para('El brillo es la media del gris OpenCV, Ygris = 0,114B + 0,587G + 0,299R, en escala de 0 a 255. El contraste es su desviación estándar. La nitidez se aproxima mediante la varianza del Laplaciano con ksize = 1 y la energía media de gradientes Sobel de tamaño 3, denominada Tenengrad. Se calculan sin redimensionar las fotografías.')
para('El ruido, la textura y la escala también afectan estos indicadores; se interpretan junto con la evidencia visual.')

page('Interpretación y selección del problema')
figure(7,'Distribución tonal y respuesta de bordes por condición',FIG/'07_diagnostico.png',5.8,'El panel derecho usa las regiones normalizadas, mientras que la tabla 2 utiliza las imágenes completas.')
table(3,'Varianza del Laplaciano en el objeto normalizado',['Imagen','I1','I2','I3','I4','I5'],[['Varianza']+[fmt(r['roi_laplaciano']) for r in rows]],[1.1,1.08,1.08,1.08,1.08,1.08])
para(f"I3 registra la menor respuesta global del Laplaciano ({fmt(rows[2]['laplaciano'])}) y también la menor respuesta en la región normalizada ({fmt(rows[2]['roi_laplaciano'])}). La coincidencia con los contornos visualmente difusos justifica seleccionarla. La reducción frente a I1 es del {fmt((1-rows[2]['laplaciano']/rows[0]['laplaciano'])*100)} % en la medida global; esta diferencia combina degradación y cambios de escena.")
para(f"I2 alcanza un brillo medio de {fmt(rows[1]['brillo'])}, frente a {fmt(rows[0]['brillo'])} en I1. El gris medio del objeto también aumenta de {fmt(rows[0]['roi_brillo'])} a {fmt(rows[1]['roi_brillo'])}. La imagen disponible no evidencia una subexposición respecto a la referencia. La autoexposición o una iluminación distinta podrían explicar el resultado, pero no se verificaron esas causas.")
para('I4 presenta distractores y menor tamaño aparente de la lata. I5 modifica la perspectiva y obtiene una alta respuesta del Laplaciano en el recorte, aunque ello no demuestra mejor legibilidad. Para esas condiciones, una segmentación o una rectificación geométrica respondería mejor al problema que aplicar un filtro de enfoque indiscriminadamente.')

page('Tratamiento aplicado y parámetros')
para('Seleccionamos una máscara de enfoque sobre el canal Y del espacio YCrCb. El desenfoque atenúa transiciones entre intensidades; sumar una fracción de la diferencia entre la luminancia original y su versión suavizada refuerza los bordes que todavía permanecen. Usar únicamente Y permite conservar los canales de crominancia durante la operación. La conversión y la cuantización pueden producir pequeñas diferencias de color.')
para('La expresión utilizada es Y′ = clip[Y + a(Y − GaussianBlur(Y, σ)), 0, 255], con cantidad a = 0,8 y desviación σ = 1,2 píxeles. El tamaño del núcleo se calcula automáticamente mediante (0, 0). El recorte se mide antes de limitar los valores y se exporta el resultado en PNG para evitar una nueva compresión JPEG con pérdida.')
code('ycc = cv2.cvtColor(image, cv2.COLOR_BGR2YCrCb)\ny = ycc[:, :, 0].astype(np.float32)\nsuave = cv2.GaussianBlur(y, (0, 0), sigmaX=1.2)\nsin_recortar = y + 0.8 * (y - suave)\nycc[:, :, 0] = np.clip(sin_recortar, 0, 255).astype(np.uint8)\nresultado = cv2.cvtColor(ycc, cv2.COLOR_YCrCb2BGR)')
table(4,'Exploración de la cantidad de realce con sigma fijo',['Cantidad a','Laplaciano\nglobal','Tenengrad\nglobal','Recorte Y\n%'],[[fmt(r['cantidad']),fmt(r['laplaciano']),fmt(r['tenengrad']),f"{r['recorte_previo_pct']:.4f}".replace('.',',')] for r in DATA['parametros']],[1.1,1.8,1.8,1.3])
para('La cantidad 0,8 ofrece un realce intermedio con poco recorte de luminancia. La opción 1,2 eleva más los indicadores y casi duplica el recorte, pero un indicador mayor no prueba mayor fidelidad. Se escoge 0,8 como decisión conservadora para esta imagen, sin afirmar que sea un óptimo universal. No aplicamos ecualización global ni CLAHE porque el problema seleccionado es la definición de bordes, no un histograma que requiera expandirse (OpenCV, s. f.-a).')

page('Comparación visual antes y después')
figure(8,'Resultado del tratamiento en la imagen completa',FIG/'08_antes_despues.png',6.3,'Ambos paneles corresponden a I3; se conserva la misma resolución y el mismo encuadre.')
figure(9,'Detalle de la lata negra en la misma región de I3',FIG/'09_detalle.png',4.4,'Recorte nativo de 125 × 240 píxeles en ambos paneles, ampliado únicamente para visualizarlo. El tratamiento se realizó sobre la imagen completa.')
para('El borde de la lata y algunos trazos del logotipo se perciben más definidos. Las letras pequeñas siguen borrosas y algunas transiciones se acentúan. Se observa un realce moderado, sin evidencia de recuperación de información perdida.')

page('Evaluación numérica del resultado')
before,after=DATA['antes'],DATA['despues']
table(5,'Comparación global de I3 antes y después',['Indicador','Antes','Después','Cambio\n%'],[[label,fmt(before[key]),fmt(after[key]),fmt((after[key]/before[key]-1)*100)] for label,key in [('Brillo medio','brillo'),('Contraste','contraste'),('Varianza Laplaciano','laplaciano'),('Tenengrad','tenengrad')]], [2.3,1.3,1.3,1.1])
figure(10,'Efecto sobre el histograma y la respuesta a bordes',FIG/'10_evaluacion.png',6.3,'La distribución tonal cambia poco, mientras que el realce aumenta la respuesta del Laplaciano.')
para(f"La varianza del Laplaciano aumenta de {fmt(before['laplaciano'])} a {fmt(after['laplaciano'])}, un {fmt((after['laplaciano']/before['laplaciano']-1)*100)} %. Tenengrad aumenta un {fmt((after['tenengrad']/before['tenengrad']-1)*100)} %. El brillo disminuye solo {fmt(before['brillo']-after['brillo'])} niveles y el contraste cambia un {fmt((after['contraste']/before['contraste']-1)*100)} %. El comportamiento concuerda con un realce de bordes y un ajuste tonal pequeño.")
recorte_texto=f"{DATA['recorte_previo_pct']:.4f}".replace('.',',')
para(f"En el recorte nativo de la lata, el Laplaciano pasa de {fmt(DATA['antes_roi']['laplaciano'])} a {fmt(DATA['despues_roi']['laplaciano'])}. Estos valores no se comparan directamente con la tabla 3, que utiliza recortes normalizados. El {recorte_texto} % de valores Y se recorta antes de reconstruir el color; esta medida es distinta de la proporción de píxeles claros o negros del gris final.")

page('Limitaciones conclusiones y recomendaciones')
para('La máscara de enfoque mejora parcialmente la apariencia de los bordes presentes. La inspección de la figura 9 muestra que el texto fino continúa degradado, por lo que el tratamiento no satisface por sí solo una tarea de lectura precisa de etiquetas. Las medidas de bordes aumentan en parte por la propia definición del filtro; también podrían aumentar por ruido o halos. No se evaluó un clasificador, un detector ni un sistema OCR, y no se afirma una mejora de exactitud en esas tareas.')
para('Las cinco tomas presentan cambios de escala, distancia aparente, iluminación y perspectiva a la vez. El recorte del objeto reduce la influencia del fondo, pero la normalización modifica el muestreo y no proporciona correspondencia punto a punto. No se calcularon PSNR ni SSIM entre I1 e I3, porque no existe una imagen nítida alineada del mismo encuadre que permita atribuir esas métricas a restauración.')
para('La fotografía identificada como baja iluminación requiere una nueva captura controlada para demostrar esa condición. Se recomienda mantener cámara, distancia y encuadre fijos, reducir la luz y registrar la exposición. Suprimir el autoajuste de exposición cuando el equipo lo permita facilitaría relacionar el cambio de luz con el resultado. Una imagen artificialmente oscurecida serviría como simulación, pero no sustituiría la fotografía propia solicitada.')
para('Para el desenfoque, la mejor medida es repetir la captura apoyando la cámara y enfocando sobre la lata. Para I4 conviene reducir distractores o segmentar el objeto; para I5, documentar la perspectiva y, si la tarea lo exige, estimar una transformación geométrica. La práctica evidencia que cada intervención debe partir del problema observado y valorarse con imágenes, indicadores y limitaciones explícitas.')

page('Evidencias del proyecto en VS Code')
evidence=ROOT/'resultados/evidencias_vscode'
if (evidence/'01_codigo.jpg').exists():
    figure(11,'Script de tratamiento abierto en VS Code',evidence/'01_codigo.jpg',6.5,'Captura real de la ventana de VS Code del equipo. El archivo mostrado pertenece al mismo repositorio que produce las métricas y las figuras del informe.')
para('La estructura separa datos originales, código fuente, cuaderno, resultados e informe. La función leer utiliza imdecode y fromfile para admitir rutas con tildes en Windows. La función metricas reúne los descriptores, y realzar implementa la única intervención seleccionada. Los resultados numéricos se conservan en archivos CSV y JSON para facilitar la revisión.')
code('gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)\nbrillo = float(gray.mean())\ncontraste = float(gray.std())\nnitidez = float(cv2.Laplacian(gray, cv2.CV_64F).var())')

page('Resultados revisables y reproducibilidad')
if (evidence/'02_resultados.jpg').exists():
    figure(12,'Comparación generada y abierta en VS Code',evidence/'02_resultados.jpg',6.5,'Captura real del visor de imágenes de VS Code. La comparación procede del tratamiento ejecutado, no de una recreación del editor.')
para(f"La ejecución local utilizó Python {DATA['python']}, OpenCV {DATA['opencv']}, NumPy {DATA['numpy']} y Matplotlib {DATA['matplotlib']}. El procesamiento conserva los originales y utiliza parámetros fijos. El archivo manifiesto.json registra las huellas SHA-256 y las coordenadas de cada región del objeto.")
para('Repositorio y código de la práctica',False,True)
para('https://github.com/Camilo19P8/TRATAMIENTO_DE_IMAGENES-',False,True)
para('Cuaderno preparado para Google Colab',False,True)
para('https://colab.research.google.com/github/Camilo19P8/TRATAMIENTO_DE_IMAGENES-/blob/main/notebooks/Practica_tratamiento_imagenes.ipynb',False,True)
para('Para reproducir, instalar requirements.txt y ejecutar python src/practica.py desde la raíz del proyecto. El README contiene instrucciones para VS Code y Colab. La ejecución documentada se realizó localmente; el enlace a Colab ofrece el entorno alternativo solicitado para repetirla.',False,True)

page('Referencias y correspondencia con la actividad')
for author,title,container,url in [
('American Psychological Association. (s. f.). ','Sample student paper annotated','', 'https://apastyle.apa.org/style-grammar-guidelines/paper-format/student-annotated.pdf'),
('OpenCV. (s. f.-a). ','Histograms - 2: Histogram equalization','OpenCV 4.13.0 documentation. ', 'https://docs.opencv.org/4.13.0/d5/daf/tutorial_py_histogram_equalization.html'),
('OpenCV. (s. f.-b). ','Smoothing images','OpenCV 4.13.0 documentation. ', 'https://docs.opencv.org/4.13.0/d4/d13/tutorial_py_filtering.html')]:
    p=para(author,False);p.add_run(title).italic=True;p.add_run('. '+container);hyperlink(p,url)
    p.paragraph_format.left_indent=Inches(.5);p.paragraph_format.first_line_indent=Inches(-.5)
para('El filtrado gaussiano utilizado se apoya en las funciones descritas por OpenCV (s. f.-b). El formato académico sigue el modelo de trabajo estudiantil de APA. Se conservan divisiones internas visibles en las tablas por solicitud expresa de presentación; esta adaptación se aparta de la convención APA de utilizar pocas líneas horizontales y evitar líneas verticales.',False)
table(6,'Localización de los elementos solicitados',['Elemento','Evidencia en el informe'],[
['Cinco fotografías propias','Figuras 1 a 5; originales en datos/originales'],
['Diagnóstico y tabla comparativa','Tablas 1 a 3; figura 7; diagnostico.csv'],
['Problema y tratamiento justificado','Selección de I3; tabla 4; función realzar'],
['Comparación visual y numérica','Figuras 8 a 10; tabla 5; antes_despues.csv'],
['Código y enlace funcional','Figuras 11 y 12; repositorio y cuaderno'],
['Análisis conclusiones y límites','Interpretación; evaluación; recomendaciones'],
['PDF de al menos dos páginas','Informe completo con portada y referencias'],
['Condición de baja iluminación','Discrepancia documentada; requiere recaptura controlada']
],[2.3,3.7])
(ROOT/'informe').mkdir(exist_ok=True)
DOC.save(ROOT/'informe/Informe_tratamiento_imagenes.docx')
print('Informe DOCX generado')
