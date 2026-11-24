"""Matriz de análisis de interesados (poder / interés) del proyecto BCAT-YOGUR.

Las puntuaciones (de 0 a 10) están en puntuaciones_interesados.csv. Son solo para
situar los puntos en la gráfica; en el documento solo aparece bajo, medio o alto.
Para mover un interesado se cambia su interés o su poder en el CSV y se vuelve a
ejecutar el script.
"""
import csv
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

CSV = Path(__file__).with_name("puntuaciones_interesados.csv")


def leer_datos():
    """Devuelve una lista de (nº, nombre, interno/externo, interés, poder)."""
    with open(CSV, encoding="utf-8") as f:
        return [(int(r["nº"]), r["interesado"], r["tipo"], float(r["interés"]), float(r["poder"]))
                for r in csv.DictReader(f, delimiter=";")]


DATOS = leer_datos()

CUADRANTES = {  # (x0, y0): (texto, color de fondo)
    (5, 5): ("GESTIONAR\nACTIVAMENTE", "#D9D9D9"),
    (0, 5): ("MANTENER\nSATISFECHOS", "#EBEBEB"),
    (5, 0): ("MANTENER\nINFORMADOS", "#EBEBEB"),
    (0, 0): ("SUPERVISAR", "#F7F7F7"),
}


def estrategia(interes, poder):
    if poder > 5:
        return "Gestionar activamente" if interes > 5 else "Mantener satisfechos"
    return "Mantener informados" if interes > 5 else "Supervisar"


def nivel(valor):
    """Nivel que se pone en el documento: bajo (<4), medio (4 a <7) o alto (>=7)."""
    return "Bajo" if valor < 4 else "Medio" if valor < 7 else "Alto"


def dibujar(salida):
    plt.rcParams["font.family"] = "Arial"
    fig = plt.figure(figsize=(10, 6.2), dpi=200)
    ax = fig.add_axes([0.07, 0.09, 0.53, 0.86])

    for (x0, y0), (texto, color) in CUADRANTES.items():
        ancho, alto = (5.5 if x0 else 5), (5.5 if y0 else 5)
        ax.add_patch(Rectangle((x0, y0), ancho, alto, facecolor=color, edgecolor="white", lw=2, zorder=0))
        ax.text(x0 + 0.2, y0 + alto - 0.2, texto, ha="left", va="top", fontsize=7.5,
                fontweight="bold", color="#7F7F7F", zorder=1)

    for n, _, tipo, i, p in DATOS:
        interno = tipo == "Interno"
        ax.scatter(i, p, s=230, zorder=3, linewidths=1.2, edgecolors="#262626",
                   facecolors="#262626" if interno else "white")
        ax.text(i, p, str(n), ha="center", va="center", fontsize=7, fontweight="bold",
                color="white" if interno else "#262626", zorder=4)

    ax.set_xlim(0, 10.5)
    ax.set_ylim(0, 10.5)
    # Sin escala numérica: en el documento solo se habla de bajo, medio o alto
    ax.set_xticks([])
    ax.set_yticks([])
    for lado in ("top", "right"):
        ax.spines[lado].set_visible(False)
    for lado in ("left", "bottom"):
        ax.spines[lado].set_color("#7F7F7F")
    ax.set_xlabel("Nivel de interés  →", fontsize=9, fontweight="bold")
    ax.set_ylabel("Poder / autoridad  →", fontsize=9, fontweight="bold")
    ax.set_aspect("equal")

    # Leyenda numerada a la derecha
    ley = fig.add_axes([0.63, 0.09, 0.36, 0.86])
    ley.axis("off")
    y = 0.99
    for n, nombre, tipo, *_ in DATOS:
        interno = tipo == "Interno"
        ley.scatter(0.03, y, s=90, linewidths=1, edgecolors="#262626",
                    facecolors="#262626" if interno else "white", clip_on=False)
        ley.text(0.03, y, str(n), ha="center", va="center", fontsize=5.5, fontweight="bold",
                 color="white" if interno else "#262626")
        ley.text(0.09, y, nombre, va="center", fontsize=7.5, color="#262626")
        y -= 0.043
    y -= 0.02
    for texto, relleno in (("Interesado interno", "#262626"), ("Interesado externo", "white")):
        ley.scatter(0.03, y, s=90, linewidths=1, edgecolors="#262626", facecolors=relleno, clip_on=False)
        ley.text(0.09, y, texto, va="center", fontsize=7.5, style="italic", color="#595959")
        y -= 0.043
    ley.set_xlim(0, 1)
    ley.set_ylim(0, 1)

    fig.savefig(salida, dpi=200, facecolor="white")
    plt.close(fig)


if __name__ == "__main__":
    salida = Path(__file__).with_name("matriz_interesados.png")
    dibujar(salida)
    print("Matriz guardada en", salida)
    for n, nombre, _, i, p in DATOS:
        print(f"{n:>2}. {nombre:<36} interés {nivel(i):<5} poder {nivel(p):<5} {estrategia(i, p)}")
