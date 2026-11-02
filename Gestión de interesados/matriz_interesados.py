"""Matriz de análisis de interesados (poder / interés) del proyecto BCAT-YOGUR.

Para cambiar la posición de un interesado basta con tocar su interés o su poder
(escala de 1 a 10) en la lista DATOS y volver a ejecutar el script.
"""
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

# (nº, nombre corto, interno/externo, interés, poder)
DATOS = [
    (1, "Dirección General (promotor)", "Interno", 9, 10),
    (2, "Director del proyecto", "Interno", 10, 7),
    (3, "Producción", "Interno", 8, 7),
    (4, "Mantenimiento", "Interno", 8, 4),
    (5, "Operarios de planta", "Interno", 6, 2),
    (6, "Ingeniería de frío industrial", "Externo", 9, 4),
    (7, "Fabricante de las bombas de calor", "Externo", 8, 6),
    (8, "Empresa instaladora", "Externo", 7, 3),
    (9, "Servicio Territorial de Industria", "Externo", 2, 9),
    (10, "Clientes (cadenas de distribución)", "Externo", 3, 6),
    (11, "Vecinos y empresas del polígono", "Externo", 3, 2),
    (12, "Organismo de control (OCA)", "Externo", 2, 4),
    (13, "Compañía eléctrica", "Externo", 2, 3),
]

CUADRANTES = {  # (x0, y0): (texto, color de fondo)
    (5, 5): ("GESTIONAR\nACTIVAMENTE", "#D9D9D9"),
    (0, 5): ("MANTENER\nSATISFECHOS", "#EBEBEB"),
    (5, 0): ("MANTENER\nINFORMADOS", "#EBEBEB"),
    (0, 0): ("SUPERVISAR CON EL\nMÍNIMO ESFUERZO", "#F7F7F7"),
}


def estrategia(interes, poder):
    if poder > 5:
        return "Gestionar activamente" if interes > 5 else "Mantener satisfechos"
    return "Mantener informados" if interes > 5 else "Supervisar con el mínimo esfuerzo"


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
    ax.set_xticks(range(0, 11))
    ax.set_yticks(range(0, 11))
    ax.tick_params(labelsize=7, colors="#595959", length=0)
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
