"""Evidencias visuales de las etapas intermedias y sus resultados.

Ejecutar después de practica.py: python src/proceso_visual.py
No modifica fotografías originales ni cambia el tratamiento seleccionado.
"""
import cv2
import numpy as np
import matplotlib.pyplot as plt
from practica import ROOT, FIG, CASOS, leer, roi, realzar, mostrar, figura


def escala_gris(ax, image, title, vmin=0, vmax=255):
    ax.imshow(image, cmap="gray", vmin=vmin, vmax=vmax)
    ax.set_title(title, fontsize=11)
    ax.axis("off")


def main():
    FIG.mkdir(parents=True, exist_ok=True)
    original=leer(ROOT/"datos/originales/03_desenfoque.jpeg")
    tratada,_=realzar(original)
    gray=cv2.cvtColor(original,cv2.COLOR_BGR2GRAY)
    ycc=cv2.cvtColor(original,cv2.COLOR_BGR2YCrCb)
    y=ycc[:,:,0].astype(np.float32)
    suavizada=cv2.GaussianBlur(y,(0,0),sigmaX=1.2)
    residual=y-suavizada
    box=CASOS[2][2]
    x1,y1,x2,y2=box
    region=(slice(y1,y2),slice(x1,x2))

    fig,axes=plt.subplots(1,3,figsize=(12,4))
    mostrar(axes[0],original,"I3 original en color")
    escala_gris(axes[1],gray,"Gris usado para medir")
    escala_gris(axes[2],y,"Luminancia Y usada para tratar")
    figura("11_color_gris_luminancia.png")

    fig,axes=plt.subplots(1,3,figsize=(9,5))
    escala_gris(axes[0],y[region],"Luminancia original")
    escala_gris(axes[1],suavizada[region],"Suavizado gaussiano · σ=1,2")
    # Visualización del residual: cero = gris medio, misma ganancia en todo el panel.
    escala_gris(axes[2],np.clip(128+6*residual[region],0,255),"Residual · 128 + 6(Y − suave)")
    figura("12_etapas_luminancia.png")

    before=roi(original,box)
    after=roi(tratada,box)
    gb=cv2.cvtColor(before,cv2.COLOR_BGR2GRAY)
    ga=cv2.cvtColor(after,cv2.COLOR_BGR2GRAY)
    lb=np.abs(cv2.Laplacian(gb,cv2.CV_64F))
    la=np.abs(cv2.Laplacian(ga,cv2.CV_64F))
    limit=max(float(lb.max()),float(la.max()))
    fig,axes=plt.subplots(1,2,figsize=(8,5))
    escala_gris(axes[0],lb,"|Laplaciano| antes",vmax=limit)
    escala_gris(axes[1],la,"|Laplaciano| después",vmax=limit)
    figura("13_respuesta_bordes.png")

    fig,axes=plt.subplots(1,4,figsize=(11,5))
    mostrar(axes[0],before,"Original")
    for ax,amount in zip(axes[1:],(.4,.8,1.2)):
        result,_=realzar(original,amount)
        mostrar(ax,roi(result,box),f"Cantidad a={amount:.1f}")
    figura("14_parametros_detalle.png")

    # Regiones nativas de I3: logotipo de la lata y sobre de la mesa.
    details=[((395,385,510,540),"Logotipo y texto de la lata"),
             ((865,590,1000,785),"Texto y contornos del sobre")]
    fig,axes=plt.subplots(2,2,figsize=(8,7))
    for index,(detail,title) in enumerate(details):
        mostrar(axes[index,0],roi(original,detail),title+" · antes")
        mostrar(axes[index,1],roi(tratada,detail),title+" · después")
    figura("15_detalles_resultado.png")

    fig,axes=plt.subplots(1,2,figsize=(11,3.8))
    for image,label in [(gb,"Antes"),(ga,"Después")]:
        hist=cv2.calcHist([image],[0],None,[256],[0,256]).ravel()
        axes[0].plot(hist/image.size,label=label)
    axes[0].set(title="Histograma en la misma región nativa",xlabel="Gris",ylabel="Proporción")
    axes[0].legend()
    delta=ga.astype(np.int16)-gb.astype(np.int16)
    values,counts=np.unique(delta,return_counts=True)
    axes[1].bar(values,counts/delta.size,color="#315b75")
    axes[1].set(title="Cambio de intensidad por píxel",xlabel="Después − antes",ylabel="Proporción")
    plt.tight_layout()
    figura("16_cambios_region.png")
    print("Se generaron seis figuras adicionales del proceso y sus resultados.")


if __name__=="__main__":
    main()
