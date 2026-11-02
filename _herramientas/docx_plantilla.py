"""Funciones comunes para generar documentos del proyecto a partir de la plantilla.

Uso típico (ver EDT/generar_edt.py):

    from docx_plantilla import Documento
    d = Documento(PLANTILLA, titulo="GESTIÓN DEL ALCANCE", subtitulo="...")
    d.titulo("1. Requisitos", 1)
    d.par("Texto normal")
    d.tabla([2, 13], filas, cabecera=["Id", "Texto"])
    d.pie("Tabla", "Título de la tabla")
    d.guardar(SALIDA)

Detalles a tener en cuenta:
- La plantilla está guardada con ONLYOFFICE: los nombres de estilo ("Heading 1"...)
  no los encuentra doc.styles["..."], por eso se buscan recorriendo los estilos.
- ONLYOFFICE guarda algunos elementos de tabla fuera del orden del esquema de Word;
  aquí se reordenan para que el archivo valide.
"""
import copy
from datetime import date

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

GRIS_OSCURO = "404040"
GRIS_CLARO = "F2F2F2"

_ORDEN_TBL = ["jc", "tblBorders", "tblLayout", "tblCellMar", "tblLook"]
_ORDEN_LADOS = ["top", "left", "bottom", "right", "insideH", "insideV"]


def _cambiar_texto(parrafos, viejo, nuevo):
    for p in parrafos:
        for r in p.runs:
            if viejo in r.text:
                r.text = r.text.replace(viejo, nuevo)


class Documento:
    def __init__(self, plantilla, titulo, subtitulo, fecha=None):
        self.doc = Document(plantilla)
        self.estilos = {s.name: s for s in self.doc.styles}
        sec = self.doc.sections[0]
        self.ancho_cm = (sec.page_width - sec.left_margin - sec.right_margin) / 360000
        self._contadores = {}

        # Formato de tabla de la plantilla (tabla tipo ficha), con el orden corregido
        tblpr = copy.deepcopy(self.doc.tables[1]._tbl.tblPr)
        for el in tblpr.findall(qn("w:tblW")) + tblpr.findall(qn("w:tblStyle")):
            tblpr.remove(el)
        for padre, orden in [(tblpr, _ORDEN_TBL)] + [(tblpr.find(qn(f"w:{t}")), _ORDEN_LADOS)
                                                     for t in ("tblBorders", "tblCellMar")]:
            for e in sorted(padre, key=lambda e: orden.index(e.tag.split("}")[1])):
                padre.remove(e)
                padre.append(e)
        self._tblpr = tblpr

        # Se deja la portada y el índice y se borra el contenido de ejemplo
        body = self.doc.element.body
        hijos = list(body.iterchildren())
        i_indice = next(i for i, el in enumerate(hijos)
                        if el.tag == qn("w:p") and "Índice" in "".join(el.itertext()))
        corte = next(i for i in range(i_indice, len(hijos)) if 'w:type="page"' in hijos[i].xml)
        for el in hijos[corte + 1:]:
            if el.tag != qn("w:sectPr"):
                body.remove(el)

        # Portada, encabezado y pie
        ps = self.doc.paragraphs
        _cambiar_texto(ps, "[TÍTULO DEL DOCUMENTO]", titulo)
        _cambiar_texto(ps, "[Subtítulo o tipo de documento: acta de constitución, plan de gestión, informe…]",
                       subtitulo)
        _cambiar_texto(ps, "[Nombre del proyecto] ([ACRÓNIMO])",
                       "Sustitución de resistencias eléctricas por bombas de calor de CO₂ "
                       "en la fabricación de yogur (BCAT-YOGUR)")
        _cambiar_texto(ps, "[dd/mm/aaaa]", (fecha or date.today()).strftime("%d/%m/%Y"))
        for h in (sec.header, sec.first_page_header):
            _cambiar_texto(h.paragraphs, "[TÍTULO DEL DOCUMENTO]", titulo)
        _cambiar_texto(sec.footer.paragraphs, "[Acrónimo del proyecto]", "BCAT-YOGUR")
        self.doc.core_properties.title = f"{titulo.capitalize()} BCAT-YOGUR"

    # ---------- texto ----------
    def estilo(self, nombre):
        return self.estilos[nombre]

    def par(self, texto="", estilo="Normal"):
        p = self.doc.add_paragraph(style=self.estilo(estilo))
        if texto:
            p.add_run(texto)
        return p

    def titulo(self, texto, nivel):
        return self.par(texto, f"Heading {nivel}")

    def lista(self, elementos, numerada=False):
        for t in elementos:
            self.par(t, "List Number" if numerada else "List Bullet")

    def lista_negrita(self, elementos):
        """Viñetas con el principio en negrita: [("Negrita", " resto del texto"), ...]"""
        for negrita, resto in elementos:
            p = self.par("", "List Bullet")
            p.add_run(negrita).bold = True
            p.add_run(resto)

    def imagen(self, ruta, ancho_cm=None):
        p = self.par()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.keep_with_next = True
        p.add_run().add_picture(str(ruta), width=Cm(ancho_cm or self.ancho_cm))
        return p

    def pie(self, tipo, texto):
        """Título numerado de tabla o figura (campo SEQ de Word). tipo: "Tabla" o "Figura"."""
        n = self._contadores[tipo] = self._contadores.get(tipo, 0) + 1
        p = self.par("", "Caption")
        p.add_run(f"{tipo} ")
        for tag, attr in (("w:fldChar", "begin"), ("w:instrText", None), ("w:fldChar", "separate")):
            el = OxmlElement(tag)
            if attr:
                el.set(qn("w:fldCharType"), attr)
            else:
                el.set(qn("xml:space"), "preserve")
                el.text = f" SEQ {tipo} \\* ARABIC "
            p.add_run()._r.append(el)
        p.add_run(str(n))
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), "end")
        p.add_run()._r.append(el)
        p.add_run(f". {texto}")
        return p

    # ---------- tablas ----------
    @staticmethod
    def _sombrear(celda, fill):
        tcPr = celda._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
        previos = [el for el in tcPr if el.tag in (qn("w:cnfStyle"), qn("w:tcW"), qn("w:gridSpan"),
                                                    qn("w:vMerge"), qn("w:tcBorders"))]
        if previos:
            previos[-1].addnext(shd)
        else:
            tcPr.insert(0, shd)

    def tabla(self, anchos, filas, cabecera=None, etiquetas=False, juntas=False,
              alinear=None, negrita_filas=()):
        """Tabla con el formato de la plantilla.

        anchos: anchos de columna en cm (deben sumar como mucho self.ancho_cm).
        etiquetas: primera columna en gris claro y negrita (tipo ficha).
        juntas: intenta que la tabla no se parta entre páginas.
        alinear: lista con "izq"/"centro"/"der" por columna.
        negrita_filas: índices de filas (sin contar cabecera) en gris claro y negrita (totales).
        """
        al = {"izq": WD_ALIGN_PARAGRAPH.LEFT, "centro": WD_ALIGN_PARAGRAPH.CENTER,
              "der": WD_ALIGN_PARAGRAPH.RIGHT}
        n = len(filas) + (1 if cabecera else 0)
        t = self.doc.add_table(rows=n, cols=len(anchos))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        t._tbl.remove(t._tbl.tblPr)
        t._tbl.insert(0, copy.deepcopy(self._tblpr))
        for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), anchos):
            gc.set(qn("w:w"), str(int(w * 567)))
        todas = ([cabecera] if cabecera else []) + list(filas)
        for r, (fila, valores) in enumerate(zip(t.rows, todas)):
            es_cab = cabecera is not None and r == 0
            i_dato = r - (1 if cabecera else 0)
            trPr = fila._tr.get_or_add_trPr()
            cs = OxmlElement("w:cantSplit")
            cs.set(qn("w:val"), "true")
            trPr.insert(0, cs)
            if es_cab:
                h = OxmlElement("w:tblHeader")
                h.set(qn("w:val"), "true")
                trPr.append(h)
            for c, (celda, v, w) in enumerate(zip(fila.cells, valores, anchos)):
                celda.width = Cm(w)
                celda.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                p = celda.paragraphs[0]
                total = not es_cab and i_dato in negrita_filas
                if es_cab:
                    p.style = self.estilo("Cabecera tabla")
                    self._sombrear(celda, GRIS_OSCURO)
                else:
                    p.style = self.estilo("Texto tabla")
                    if (etiquetas and c == 0) or total:
                        self._sombrear(celda, GRIS_CLARO)
                if alinear:
                    p.alignment = al[alinear[c]]
                run = p.add_run(str(v))
                if not es_cab and ((etiquetas and c == 0) or total):
                    run.bold = True
                if juntas and r < n - 1:
                    p.paragraph_format.keep_with_next = True
        return t

    def guardar(self, salida):
        self.doc.save(str(salida))
