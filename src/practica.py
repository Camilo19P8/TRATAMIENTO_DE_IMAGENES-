"""Diagnóstico de cinco fotografías propias y realce moderado de bordes.

Ejecutar desde la raíz: python src/practica.py
Las métricas son descriptores; no equivalen a calidad perceptual ni a recuperación.
"""
from pathlib import Path
import csv
import hashlib
import json
import platform

import cv2
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "resultados"
FIG = OUT / "figuras"
# Regiones manuales de la misma lata negra Monster, en píxeles x1,y1,x2,y2.
CASOS = [
    ("01_referencia.jpeg", "Referencia", (440, 205, 585, 530)),
    ("02_iluminacion.jpeg", "Iluminación declarada baja", (440, 170, 580, 485)),
    ("03_desenfoque.jpeg", "Desenfoque o movimiento", (385, 300, 510, 540)),
    ("04_fondo_complejo.jpeg", "Fondo complejo", (335, 180, 440, 385)),
    ("05_angulo.jpeg", "Cambio de ángulo", (475, 340, 610, 565)),
]


def leer(path):
    """Lectura compatible con rutas Unicode en Windows."""
    image = cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)
    if image is None:
        raise ValueError(f"No se pudo leer {path}")
    return image


def guardar(path, image):
    ok, encoded = cv2.imencode(path.suffix, image)
    if not ok:
        raise ValueError(f"No se pudo codificar {path}")
    encoded.tofile(path)


def metricas(image):
    """Gris OpenCV: 0.114 B + 0.587 G + 0.299 R, rango 0 a 255."""
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gx = cv2.Sobel(gray, cv2.CV_64F, 1, 0, ksize=3)
    gy = cv2.Sobel(gray, cv2.CV_64F, 0, 1, ksize=3)
    return {
        "brillo": float(gray.mean()),
        "contraste": float(gray.std()),
        "laplaciano": float(cv2.Laplacian(gray, cv2.CV_64F).var()),
        "tenengrad": float(np.mean(gx * gx + gy * gy)),
        "oscuros_pct": float(np.mean(gray <= 5) * 100),
        "claros_pct": float(np.mean(gray >= 250) * 100),
    }


def roi(image, box, normalizar=False):
    x1, y1, x2, y2 = box
    crop = image[y1:y2, x1:x2]
    if normalizar:
        crop = cv2.resize(crop, (128, 256), interpolation=cv2.INTER_AREA)
    return crop


def realzar(image, amount=0.8, sigma=1.2):
    """Máscara de enfoque solo en luminancia; no es deconvolución."""
    ycc = cv2.cvtColor(image, cv2.COLOR_BGR2YCrCb)
    y = ycc[:, :, 0].astype(np.float32)
    suave = cv2.GaussianBlur(y, (0, 0), sigmaX=sigma)
    sin_recortar = y + amount * (y - suave)
    recorte = float(np.mean((sin_recortar < 0) | (sin_recortar > 255)) * 100)
    ycc[:, :, 0] = np.clip(sin_recortar, 0, 255).astype(np.uint8)
    return cv2.cvtColor(ycc, cv2.COLOR_YCrCb2BGR), recorte


def tabla(path, rows):
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)


def mostrar(ax, image, title):
    ax.imshow(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
    ax.set_title(title, fontsize=12)
    ax.axis("off")


def figura(name):
    plt.savefig(FIG / name, dpi=190, bbox_inches="tight", facecolor="white")
    plt.close()


def main():
    FIG.mkdir(parents=True, exist_ok=True)
    images, rows, manifest = [], [], []
    for name, condition, box in CASOS:
        path = ROOT / "datos" / "originales" / name
        image = leer(path)
        images.append(image)
        h, w, c = image.shape
        with Image.open(path) as pil:
            fmt = pil.format
        row = {"archivo": name, "condicion": condition, "ancho": w,
               "alto": h, "canales": c, "formato": fmt,
               "bytes": path.stat().st_size, **metricas(image)}
        row.update({"roi_" + k: v for k, v in metricas(roi(image, box, True)).items()})
        rows.append(row)
        manifest.append({"archivo": name, "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                         "roi_xyxy": box})
        index = len(images)
        fig, ax = plt.subplots(figsize=(11, 6.2))
        mostrar(ax, image, f"I{index} · {condition}")
        x1, y1, x2, y2 = box
        ax.add_patch(plt.Rectangle((x1, y1), x2-x1, y2-y1, fill=False,
                                  edgecolor="#ffca28", linewidth=2))
        figura(f"{index:02d}_captura.png")
    tabla(OUT / "diagnostico.csv", rows)
    fig, axes = plt.subplots(1, 5, figsize=(12, 5))
    for ax, image, case in zip(axes, images, CASOS):
        mostrar(ax, roi(image, case[2], True), case[0][:2])
    figura("06_mismo_objeto.png")
    fig, axes = plt.subplots(1, 2, figsize=(11, 3.7))
    for image, case in zip(images, CASOS):
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        hist = cv2.calcHist([gray], [0], None, [256], [0, 256]).ravel()
        axes[0].plot(hist / gray.size, label=case[0][:2], linewidth=1)
    axes[0].set(xlabel="Intensidad gris", ylabel="Proporción", title="Histogramas globales")
    axes[0].legend()
    axes[1].bar(["I1", "I2", "I3", "I4", "I5"], [r["roi_laplaciano"] for r in rows], color="#315b75")
    axes[1].set(title="Varianza del Laplaciano en ROI normalizada", ylabel="Varianza")
    plt.tight_layout()
    figura("07_diagnostico.png")
    original = images[2]
    candidates = []
    for amount in (0.4, 0.8, 1.2):
        processed, clipping = realzar(original, amount)
        candidates.append({"cantidad": amount, "sigma": 1.2, "recorte_previo_pct": clipping,
                           **metricas(processed),
                           **{"roi_"+k: v for k,v in metricas(roi(processed, CASOS[2][2])).items()}})
    tabla(OUT / "parametros.csv", candidates)
    processed, clipping = realzar(original)
    guardar(OUT / "03_desenfoque_tratada.png", processed)
    before, after = metricas(original), metricas(processed)
    tabla(OUT / "antes_despues.csv", [{"estado": "Antes", **before}, {"estado": "Después", **after}])
    fig, axes = plt.subplots(1, 2, figsize=(12, 4.2))
    mostrar(axes[0], original, "Antes · I3 original")
    mostrar(axes[1], processed, "Después · máscara de enfoque moderada")
    figura("08_antes_despues.png")
    fig, axes = plt.subplots(1, 2, figsize=(8, 5))
    mostrar(axes[0], roi(original, CASOS[2][2]), "Antes · misma región nativa")
    mostrar(axes[1], roi(processed, CASOS[2][2]), "Después · misma región nativa")
    figura("09_detalle.png")
    fig, axes = plt.subplots(1, 2, figsize=(10, 3.5))
    for img, label in [(original, "Antes"), (processed, "Después")]:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        hist = cv2.calcHist([gray], [0], None, [256], [0, 256]).ravel()
        axes[0].plot(hist/gray.size, label=label)
    axes[0].set(xlabel="Intensidad gris", ylabel="Proporción", title="Distribución tonal")
    axes[0].legend()
    axes[1].bar(["Antes", "Después"], [before["laplaciano"], after["laplaciano"]], color=["#a2a9ae", "#315b75"])
    axes[1].set(title="Respuesta a bordes", ylabel="Varianza del Laplaciano")
    plt.tight_layout()
    figura("10_evaluacion.png")
    summary = {"diagnostico": rows, "antes": before, "despues": after,
               "antes_roi": metricas(roi(original, CASOS[2][2])),
               "despues_roi": metricas(roi(processed, CASOS[2][2])),
               "recorte_previo_pct": clipping, "parametros": candidates,
               "python": platform.python_version(), "opencv": cv2.__version__,
               "numpy": np.__version__, "matplotlib": matplotlib.__version__}
    (OUT / "resumen.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    (OUT / "manifiesto.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print("Cinco imágenes analizadas. Resultados escritos en resultados/.")
    for row in rows:
        print(f"{row['archivo']}: brillo={row['brillo']:.2f}; contraste={row['contraste']:.2f}; Laplaciano={row['laplaciano']:.2f}")
    print(f"I3: Laplaciano {before['laplaciano']:.2f} -> {after['laplaciano']:.2f}")


if __name__ == "__main__":
    main()
