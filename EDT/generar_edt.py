"""Genera el documento de gestión del alcance (EDT) del proyecto BCAT-YOGUR.

Crea el diagrama de la EDT (edt_diagrama.png) y el documento
Gestion_alcance_EDT_BCAT-YOGUR.docx a partir de la plantilla.

Ejecutar desde cualquier sitio:  python generar_edt.py
Ojo: si el .docx ya se ha retocado a mano, volver a ejecutar el script lo sobrescribe.
"""
import sys
import textwrap
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

AQUI = Path(__file__).resolve().parent
RAIZ = AQUI.parent
sys.path.insert(0, str(RAIZ / "_herramientas"))
from docx_plantilla import Documento  # noqa: E402

PLANTILLA = RAIZ / "Plantilla_documentos_BCAT-YOGUR.docx"
DIAGRAMA = AQUI / "edt_diagrama.png"
SALIDA = AQUI / "Gestion_alcance_EDT_BCAT-YOGUR.docx"

# ======================= DATOS DE LA EDT =======================
# Nivel 2 (entregables): id -> (nombre, responsable, plazo)
ENTREGABLES = {
    "1": ("Gestión del proyecto", "Director del proyecto", "Nov 2026 – Mar 2028"),
    "2": ("Estudios previos", "Ingeniería de frío", "Nov 2026 – Feb 2027"),
    "3": ("Proyecto ejecutivo y legalización", "Ingeniería de frío", "Mar 2027 – Dic 2027"),
    "4": ("Instalación del sistema", "Empresa instaladora", "May 2027 – Nov 2027"),
    "5": ("Puesta en marcha y formación", "Fabricante de las bombas de calor", "Nov 2027 – Dic 2027"),
    "6": ("Verificación de resultados", "Director del proyecto", "Ene 2028 – Mar 2028"),
}

# Nivel 3 (paquetes de trabajo):
# id, nombre corto (diagrama), objetivo, descripción, responsable, hito, coste (€), tiempo
PAQUETES = [
    ("1.1", "Planificación", "Arrancar el proyecto con un plan claro",
     "Aprobar el acta de constitución, formar el equipo y preparar la planificación inicial.",
     "Director del proyecto", "H0", 3_000, "Nov 2026 (1 mes)"),
    ("1.2", "Seguimiento y control", "Controlar plazo, coste y alcance",
     "Reuniones mensuales con la Dirección, informes de avance y gestión de cambios.",
     "Director del proyecto", "—", 15_000, "Nov 2026 – Mar 2028 (17 meses)"),
    ("1.3", "Contratación y compras", "Elegir proveedores y cerrar contratos",
     "Petición de ofertas, comparación, adjudicación de la obra y pedido de los equipos.",
     "Director del proyecto", "H4", 4_000, "Abr – May 2027 (2 meses)"),
    ("1.4", "Cierre del proyecto", "Cerrar formalmente el proyecto",
     "Informe de cierre con las lecciones aprendidas y archivo de la documentación.",
     "Director del proyecto", "H8", 3_000, "Mar 2028 (1 mes)"),
    ("2.1", "Auditoría energética", "Conocer cuánto calor y frío se gasta y cuándo",
     "Medición de consumos de pasteurización, incubación, CIP y frío, y perfil de demanda horario.",
     "Ingeniería de frío industrial", "H1", 10_000, "Nov 2026 – Ene 2027 (3 meses)"),
    ("2.2", "Integración térmica y selección", "Elegir y dimensionar la bomba de calor",
     "Estudio de integración térmica (pinch) y comparación de tecnologías y fabricantes.",
     "Ingeniería de frío industrial", "H2", 8_000, "Feb 2027 (1 mes)"),
    ("3.1", "Proyecto ejecutivo", "Definir la instalación al detalle",
     "Memoria, planos, presupuesto y pliegos para pedir ofertas y montar.",
     "Ingeniería de frío industrial", "H3", 20_000, "Mar – Abr 2027 (2 meses)"),
    ("3.2", "Legalización", "Poder poner la instalación en servicio",
     "Tramitación ante el Servicio Territorial de Industria (RSIF, REBT y equipos a presión) e inspección del OCA.",
     "Ingeniería de frío industrial", "H7", 4_000, "Abr – Dic 2027 (tramitación)"),
    ("4.1", "Bombas de calor", "Disponer de los equipos a tiempo",
     "Fabricación, pruebas en fábrica y transporte de las dos bombas de calor de CO₂.",
     "Fabricante de las bombas de calor", "H6", 330_000, "May – Oct 2027 (6 meses)"),
    ("4.2", "Integración con el frío", "Aprovechar el calor de la instalación frigorífica",
     "Intercambiadores y circuito de recuperación de calor de condensación.",
     "Empresa instaladora", "H5", 60_000, "Ago – Oct 2027 (3 meses)"),
    ("4.3", "Depósitos de acumulación", "Cubrir los picos del CIP y desplazar consumo",
     "Suministro y colocación de los depósitos de 95 °C y de 45–50 °C.",
     "Empresa instaladora", "H6", 55_000, "Sep – Nov 2027 (3 meses)"),
    ("4.4", "Red hidráulica", "Llevar el calor a cada consumidor",
     "Tuberías por niveles de temperatura, bombas con variador y valvulería.",
     "Empresa instaladora", "H5", 70_000, "Jul – Sep 2027 (3 meses)"),
    ("4.5", "Adaptación de consumidores", "Que pasteurizador, incubación y CIP usen el nuevo calor",
     "Cambios en el pasteurizador, los tanques de incubación y el CIP durante la parada de agosto.",
     "Empresa instaladora", "H5", 80_000, "Ago 2027 (parada, 1 mes)"),
    ("4.6", "Electricidad y control", "Controlar y medir el sistema",
     "Cuadros eléctricos, PLC/SCADA y contadores de energía térmica y eléctrica.",
     "Empresa instaladora", "H6", 50_000, "Sep – Nov 2027 (3 meses)"),
    ("4.7", "Montaje de las bombas de calor", "Dejar los equipos principales instalados",
     "Colocación, conexión hidráulica y frigorífica de las bombas de calor.",
     "Empresa instaladora", "H6", 12_000, "Oct – Nov 2027 (2 meses)"),
    ("5.1", "Pruebas y recepción", "Comprobar que todo funciona como se pidió",
     "Pruebas en vacío y en carga, ajuste del control y firma del acta de recepción.",
     "Fabricante de las bombas de calor", "H7", 15_000, "Dic 2027 (1 mes)"),
    ("5.2", "Manuales y formación", "Que el personal sepa usar y mantener el sistema",
     "Entrega de manuales y formación práctica a Mantenimiento y a los operarios.",
     "Fabricante de las bombas de calor", "H7", 8_000, "Nov – Dic 2027 (2 meses)"),
    ("6.1", "Operación asistida", "Acompañar los primeros meses de uso",
     "Ajustes y resolución de incidencias con apoyo del fabricante y de la ingeniería.",
     "Mantenimiento", "H8", 5_000, "Ene – Mar 2028 (3 meses)"),
    ("6.2", "Informe de ahorros", "Demostrar los ahorros conseguidos",
     "Comparación de consumos antes y después con los datos de los contadores.",
     "Ingeniería de frío industrial", "H8", 3_000, "Mar 2028 (1 mes)"),
]
CONTINGENCIA = 75_500
assert sum(p[6] for p in PAQUETES) + CONTINGENCIA == 830_500, "La EDT no cuadra con el presupuesto del acta"


def euros(v):
    return f"{v:,.0f} €".replace(",", ".")


# ======================= DIAGRAMA =======================
def dibujar_edt(salida):
    plt.rcParams["font.family"] = "Arial"
    W, n = 5.9, len(ENTREGABLES)
    col = W / n
    hijos = {k: [p for p in PAQUETES if p[0].split(".")[0] == k] for k in ENTREGABLES}
    max_h = max(len(v) for v in hijos.values())
    alto_l1, alto_l2, alto_l3, paso = 0.42, 0.5, 0.4, 0.48
    y_l1 = 0.05
    y_l2 = y_l1 + alto_l1 + 0.3
    y_l3 = y_l2 + alto_l2 + 0.12
    H = y_l3 + max_h * paso + 0.05

    fig = plt.figure(figsize=(W, H), dpi=250)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, W)
    ax.set_ylim(H, 0)
    ax.axis("off")
    linea = dict(color="#7F7F7F", lw=0.8, zorder=1)

    def caja(x, y, w, h, texto, fondo, color_texto, size, negrita=False, ancho=17):
        ax.add_patch(Rectangle((x, y), w, h, facecolor=fondo, edgecolor="#404040", lw=0.7, zorder=2))
        texto = "\n".join(textwrap.fill(linea, ancho) for linea in texto.split("\n"))
        ax.text(x + w / 2, y + h / 2, texto, ha="center", va="center",
                fontsize=size, color=color_texto, fontweight="bold" if negrita else "normal",
                linespacing=1.15, zorder=3)

    # Nivel 1
    w1 = 2.9
    caja((W - w1) / 2, y_l1, w1, alto_l1, "BCAT-YOGUR\nSustitución de resistencias por bombas de calor",
         "#404040", "white", 7, True, ancho=60)
    y_barra = y_l1 + alto_l1 + 0.15
    ax.plot([W / 2, W / 2], [y_l1 + alto_l1, y_barra], **linea)
    ax.plot([col / 2, W - col / 2], [y_barra, y_barra], **linea)

    # Niveles 2 y 3
    w2, w3, sangria = col - 0.08, col - 0.17, 0.09
    for i, (k, (nombre, *_)) in enumerate(ENTREGABLES.items()):
        x = i * col + 0.04
        ax.plot([x + w2 / 2] * 2, [y_barra, y_l2], **linea)
        caja(x, y_l2, w2, alto_l2, f"{k}. {nombre}", "#D9D9D9", "#262626", 6.3, True, ancho=15)
        if hijos[k]:
            x_linea = x + 0.045
            y_ult = y_l3 + (len(hijos[k]) - 1) * paso + alto_l3 / 2
            ax.plot([x_linea, x_linea], [y_l2 + alto_l2, y_ult], **linea)
            for j, p in enumerate(hijos[k]):
                y = y_l3 + j * paso
                ax.plot([x_linea, x + sangria], [y + alto_l3 / 2] * 2, **linea)
                caja(x + sangria, y, w3, alto_l3, f"{p[0]} {p[1]}", "white", "#262626", 5.8, ancho=18)

    fig.savefig(salida, dpi=250, facecolor="white")
    plt.close(fig)


# ======================= DOCUMENTO =======================
def generar_documento():
    d = Documento(PLANTILLA, "GESTIÓN DEL ALCANCE",
                  "Requisitos, declaración del alcance y EDT")
    A = d.ancho_cm

    # ---------- 1. Requisitos ----------
    d.titulo("1. Recopilar requisitos", 1)
    d.par("Los requisitos salen del acta de constitución y del registro de interesados. Para completarlos se "
          "han tenido en cuenta las necesidades de Producción y Mantenimiento, que son quienes van a usar el "
          "sistema, y las exigencias del Servicio Territorial de Industria para poder legalizar la instalación.")
    d.par("Cada requisito se relaciona con el objetivo del acta al que responde y con el paquete de la EDT "
          "donde se resuelve (apartado 3), y se indica cómo se va a medir y comprobar.")

    REQ = [
        ("R01", "La bomba de calor da todo el calor de pasteurización, incubación y CIP. Las resistencias quedan solo como respaldo.",
         "Alta", "Funcional", "Acta de constitución", "Alcance", "4.1, 5.1",
         "Al menos el 95 % del calor del proceso cubierto por la bomba de calor", "Lectura de contadores durante la operación asistida"),
        ("R02", "Agua caliente a 90–95 °C (pasteurizador y CIP) y a 45–50 °C (incubación), con ±1 °C en la incubación.",
         "Alta", "Técnico", "Producción", "Alcance", "4.5, 5.1",
         "Temperaturas registradas en el SCADA", "Pruebas de puesta en marcha con producto"),
        ("R03", "Rendimiento medio anual (SCOP) de 3,0 o más solo en calor, y de 5 o más contando el frío aprovechado.",
         "Alta", "Rendimiento", "Acta de constitución", "Coste", "2.2, 6.2",
         "SCOP calculado con los contadores", "Informe de ahorros"),
        ("R04", "Refrigerante natural (CO₂, R744) conforme al Reglamento (UE) 2024/573.",
         "Alta", "Normativo", "Acta de constitución", "Alcance", "2.2, 4.1",
         "Refrigerante indicado en la ficha técnica", "Revisión del proyecto ejecutivo"),
        ("R05", "Circuitos indirectos: el refrigerante nunca está en contacto con el producto.",
         "Alta", "Seguridad alimentaria", "Producción", "Alcance", "3.1, 4.5",
         "Esquema hidráulico del proyecto", "Revisión del diseño e inspección del montaje"),
        ("R06", "Depósitos de acumulación para cubrir los picos del CIP y consumir más en horas valle.",
         "Media", "Técnico", "Acta de constitución", "Coste", "4.3",
         "El CIP se hace sin bajar de 85 °C", "Prueba de un ciclo de CIP completo"),
        ("R07", "No perder producción: la línea solo se toca en las paradas programadas.",
         "Alta", "Operativo", "Producción", "Tiempo", "1.2, 4.5",
         "0 paradas no programadas por el proyecto", "Partes de producción"),
        ("R08", "Instalación legalizada según RSIF, REBT y Directiva de Equipos a Presión.",
         "Alta", "Normativo", "Servicio Territorial de Industria", "Tiempo", "3.2",
         "Inscripción de la instalación concedida", "Certificado de Industria e informe del OCA"),
        ("R09", "Manuales y formación para Mantenimiento y para los operarios.",
         "Media", "Formación", "Mantenimiento y operarios", "Alcance", "5.2",
         "Todo el personal implicado formado", "Hoja de firmas de la formación"),
    ]
    d.tabla([1.0, 7.0, 1.6, 2.4, 3.0], [r[:5] for r in REQ],
            cabecera=["Id", "Requisito", "Prioridad", "Categoría", "Fuente"])
    d.pie("Tabla", "Matriz de trazabilidad de requisitos: información de los requisitos")
    d.tabla([1.0, 1.8, 1.8, 5.2, 5.2], [(r[0],) + r[5:] for r in REQ],
            cabecera=["Id", "Objetivo", "Paquete EDT", "Métrica", "Validación"])
    d.pie("Tabla", "Matriz de trazabilidad de requisitos: relación de trazabilidad")

    # ---------- 2. Alcance ----------
    d.titulo("2. Definir el alcance", 1)
    d.titulo("2.1. Alcance del producto", 2)
    d.par("El resultado del proyecto es un sistema de bombas de calor que sustituye a las resistencias "
          "eléctricas como fuente de calor de la planta. Está formado por:")
    d.lista([
        "Dos bombas de calor de alta temperatura de CO₂ (R744) de unos 175 kW térmicos cada una.",
        "Un circuito de recuperación de calor conectado a la instalación frigorífica existente.",
        "Depósitos de acumulación de agua a 95 °C y a 45–50 °C.",
        "Una red hidráulica separada por niveles de temperatura, con bombas con variador.",
        "El pasteurizador, los tanques de incubación y el CIP adaptados para usar el nuevo calor.",
        "Un sistema de control PLC/SCADA con contadores de energía térmica y eléctrica.",
    ])
    d.titulo("2.2. Alcance del proyecto", 2)
    d.par("El proyecto incluye todo el trabajo necesario para que ese sistema funcione: desde estudiar el "
          "consumo actual de la planta hasta comprobar los ahorros después de tres meses de uso.")
    d.par("Incluye:", "Etiqueta")
    d.lista([
        "Auditoría energética y estudio de integración térmica.",
        "Proyecto ejecutivo y legalización de la instalación.",
        "Compra, montaje e integración de todos los equipos descritos en el apartado 2.1.",
        "Pruebas, puesta en marcha, manuales y formación del personal.",
        "Tres meses de operación asistida y el informe de ahorros.",
    ])
    d.par("No incluye:", "Etiqueta")
    d.lista([
        "Cambiar o ampliar la instalación frigorífica: solo se conecta a ella.",
        "Desmontar las resistencias, que se quedan como respaldo.",
        "Instalar placas solares de autoconsumo (se deja como posible proyecto futuro).",
        "Cambios en la receta del yogur o en la línea de envasado.",
        "El mantenimiento una vez acabada la garantía, que pasa a Mantenimiento.",
    ])
    d.titulo("2.3. Supuestos", 2)
    d.lista([
        "La parada de producción de agosto de 2027 dura lo suficiente para adaptar los consumidores (unas tres semanas).",
        "La instalación frigorífica tiene suficiente calor de condensación aprovechable.",
        "Hay espacio en la planta para las bombas de calor y los depósitos sin hacer obra de ampliación.",
        "El fabricante entrega las bombas de calor en un plazo de 4 a 6 meses desde el pedido.",
    ])
    d.titulo("2.4. Restricciones", 2)
    d.lista([
        "Presupuesto máximo de 830.500 € (sin IVA), con una desviación de como mucho el 10 %.",
        "Puesta en marcha antes del 31/12/2027 y cierre en marzo de 2028.",
        "Los trabajos en la línea solo se pueden hacer en paradas programadas.",
        "Hay que cumplir el RSIF, el REBT, la Directiva de Equipos a Presión y el APPCC.",
    ])

    # ---------- 3. EDT ----------
    d.titulo("3. Crear la EDT", 1)
    d.par("La EDT divide el proyecto en seis entregables (nivel 2) y cada uno en paquetes de trabajo más "
          "pequeños (nivel 3), que son los que se pueden asignar a un responsable, presupuestar y planificar.")
    d.imagen(DIAGRAMA)
    d.pie("Figura", "Estructura de desglose del trabajo (EDT). Fuente: elaboración propia")

    d.titulo("3.1. Diccionario de la EDT", 2)
    d.par("En el diccionario se describe cada paquete de trabajo. Las filas en gris son los entregables, con "
          "la suma de sus paquetes.")
    filas, grises = [], []
    for k, (nombre, resp, plazo) in ENTREGABLES.items():
        hijos = [p for p in PAQUETES if p[0].split(".")[0] == k]
        grises.append(len(filas))
        filas.append([k, nombre, "", resp, "", euros(sum(p[6] for p in hijos)), plazo])
        for pid, _, objetivo, desc, resp_p, hito, coste, tiempo in hijos:
            filas.append([pid, objetivo, desc, resp_p, hito, euros(coste), tiempo])
    total_paq = sum(p[6] for p in PAQUETES)
    grises += [len(filas), len(filas) + 1, len(filas) + 2]
    filas += [["", "Total paquetes de trabajo", "", "", "", euros(total_paq), ""],
              ["", "Reserva de contingencia (10 %)", "La gestiona el director del proyecto", "", "",
               euros(CONTINGENCIA), ""],
              ["", "TOTAL (sin IVA)", "", "", "", euros(total_paq + CONTINGENCIA), "17 meses"]]
    d.tabla([0.9, 2.9, 4.4, 2.2, 0.9, 1.7, 2.0], filas,
            cabecera=["Id", "Objetivo", "Descripción", "Responsable", "Hito", "Coste estimado", "Tiempo de ejecución"],
            alinear=["izq", "izq", "izq", "izq", "centro", "der", "izq"], negrita_filas=grises)
    d.pie("Tabla", "Diccionario de la EDT")

    # ---------- 4. Verificar ----------
    d.titulo("4. Verificar el alcance", 1)
    d.par("Verificar el alcance es aceptar formalmente cada entregable cuando se termina. No es lo mismo que el "
          "control de calidad: aquí lo que se busca es que el cliente o el promotor revisen el entregable y den "
          "su conformidad por escrito.")
    d.tabla([3.6, 3.6, 6.2, 1.6], [
        ["2. Estudios previos", "Director del proyecto y Dirección General", "Revisión de los informes de auditoría y de integración térmica, y aprobación de la tecnología elegida.", "H1, H2"],
        ["3. Proyecto ejecutivo y legalización", "Dirección General, Producción y Mantenimiento", "Aprobación firmada del proyecto ejecutivo. La legalización se da por buena con la inscripción en Industria.", "H3, H7"],
        ["4. Instalación del sistema", "Ingeniería (dirección de obra) y Mantenimiento", "Inspección del montaje con una lista de comprobación e informe del OCA de los equipos a presión.", "H5, H6"],
        ["5. Puesta en marcha y formación", "Producción y Dirección General", "Acta de recepción firmada tras comprobar temperaturas y cobertura del calor, y hoja de firmas de la formación.", "H7"],
        ["6. Verificación de resultados", "Dirección General", "Aprobación del informe de ahorros y del informe de cierre.", "H8"],
    ], cabecera=["Entregable", "Quién lo acepta", "Cómo se acepta", "Hito"], etiquetas=True)
    d.pie("Tabla", "Aceptación de los entregables")

    # ---------- 5. Controlar ----------
    d.titulo("5. Controlar el alcance", 1)
    d.par("La línea base del alcance es este documento: la declaración del alcance, la EDT y su diccionario. "
          "Cualquier cambio sobre ella sigue estos pasos:")
    d.lista([
        "Quien pide el cambio lo hace por escrito, explicando qué quiere cambiar y por qué.",
        "El director del proyecto analiza cómo afecta al coste, al plazo y a los requisitos.",
        "Si el cambio es pequeño (hasta 25.000 € por partida y sin mover los hitos H3, H7 y H8) lo aprueba "
        "el director del proyecto. Si cambia el alcance, la fecha final o desvía el coste más de un 10 %, lo "
        "tiene que aprobar la Dirección General.",
        "Si se aprueba, se actualizan la EDT, el diccionario y el presupuesto, y se avisa a los afectados.",
    ], numerada=True)
    d.par("Además, en la reunión mensual de seguimiento (paquete 1.2) se revisa el avance de cada paquete de "
          "trabajo para detectar a tiempo trabajos que no estaban previstos o que se están desviando.")

    d.guardar(SALIDA)


if __name__ == "__main__":
    dibujar_edt(DIAGRAMA)
    generar_documento()
    print("Generados:", DIAGRAMA.name, "y", SALIDA.name)
