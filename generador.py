import os
import tempfile
import sys
from io import BytesIO
from collections.abc import Iterable

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm
from docx.shared import Pt
from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.lib.utils import ImageReader, simpleSplit
from reportlab.pdfgen import canvas
import win32com.client


class Generador:

    FORMATO_PDF = 17
    FORMATO_DOCX = 12
    PUNTOS_POR_CM = 28.3464567

    # Constantes de la API de Word (win32com) usadas para el relleno con "=".
    WD_COLLAPSE_END = 0
    WD_CHARACTER = 1
    WD_FIRST_CHARACTER_LINE_NUMBER = 10
    WD_STATISTIC_LINES = 1
    TEXTO_FOOTER = (
        "Se emite el presente testimonio de conformidad con lo regulado por los artículos 24° y 28° del Decreto "
        "Legislativo N° 1049 - Decreto Legislativo del Notariado, en concordancia con lo regulado por la ley de "
        "firmas y certificados digitales y su reglamento aprobado por Decreto Supremo N° 052-2008-PCM. "
        "Este documento notarial en formato digital contiene el traslado de la Escritura Pública, cuya verificación "
        "se encuentra disponible por medio del código QR del link de verificación indicado."
    )

    def __init__(self, ruta):

        self.doc = Document(ruta)
        self.logo_temporal = None
        self._margen_inferior_minimo_cm = None
        self._margen_superior_minimo_cm = 1.0
        self._ancho_extra_cm = 0.0
        # Tuplas (indice_parrafo, caracter) de los párrafos que deben
        # rellenarse hasta el margen derecho. El relleno real se hace más
        # adelante, con Word abierto, para usar su medición exacta.
        self._indices_pendientes_iguales = []

    def _base_recursos(self):
        return getattr(sys, "_MEIPASS", os.path.dirname(__file__))

    def configurar_reserva_pie(self, margen_inferior_minimo_cm=3.2):
        self._margen_inferior_minimo_cm = margen_inferior_minimo_cm

    def configurar_acta_vehicular(self):
        self.configurar_reserva_pie(3.2)
        # Mantener el ancho util original para no provocar reflujo de texto.
        self._ancho_extra_cm = 0.0

    def _obtener_logo_suavizado(self):

        if self.logo_temporal and os.path.exists(self.logo_temporal):
            return self.logo_temporal

        logo_original = os.path.join(self._base_recursos(), "assets", "logo.png")

        if not os.path.exists(logo_original):
            return logo_original

        with Image.open(logo_original).convert("RGBA") as imagen:
            fondo = Image.new("RGBA", imagen.size, (255, 255, 255, 255))
            suavizada = Image.blend(imagen, fondo, 0.55)

            archivo_temp = tempfile.NamedTemporaryFile(
                suffix=".png",
                delete=False,
            )

            self.logo_temporal = archivo_temp.name
            archivo_temp.close()
            suavizada.save(self.logo_temporal)

        return self.logo_temporal

    def _obtener_ruta_firma(self):

        base_dir = self._base_recursos()
        candidatos = [
            os.path.join(base_dir, "firma.png"),
            os.path.join(base_dir, "assets", "firma.png"),
        ]

        for ruta in candidatos:
            if os.path.exists(ruta):
                return ruta

        return candidatos[0]

    def _estampar_firma_en_pdf(self, ruta_pdf):

        ruta_firma = self._obtener_ruta_firma()
        ruta_logo = self._obtener_logo_suavizado()

        if not os.path.exists(ruta_firma):
            return

        lector = PdfReader(ruta_pdf)
        escritor = PdfWriter()
        imagen_firma = ImageReader(ruta_firma)
        imagen_logo = ImageReader(ruta_logo) if os.path.exists(ruta_logo) else None
        total_paginas = len(lector.pages)

        for indice, pagina in enumerate(lector.pages):
            ancho_pagina = float(pagina.mediabox.width)
            alto_pagina = float(pagina.mediabox.height)
            margen_izquierdo = 2.5 * self.PUNTOS_POR_CM

            ancho_firma = 1.4 * self.PUNTOS_POR_CM
            ancho_original, alto_original = imagen_firma.getSize()
            alto_firma = ancho_firma * (alto_original / ancho_original)

            x = (margen_izquierdo - ancho_firma) / 2
            y = (alto_pagina - alto_firma) / 2

            overlay_stream = BytesIO()
            overlay = canvas.Canvas(overlay_stream, pagesize=(ancho_pagina, alto_pagina))
            overlay.drawImage(
                imagen_firma,
                x,
                y,
                width=ancho_firma,
                height=alto_firma,
                preserveAspectRatio=True,
                mask="auto",
            )

            if indice == 0 or indice == (total_paginas - 1):
                # Pie solo en primera y ultima pagina, con separacion del borde inferior.
                y_base_footer = 1.05 * self.PUNTOS_POR_CM
                ancho_bloque = 16.0 * self.PUNTOS_POR_CM
                x_bloque = ((ancho_pagina - ancho_bloque) / 2) + (0.9 * self.PUNTOS_POR_CM)

                ancho_col_logo = 4.8 * self.PUNTOS_POR_CM
                ancho_col_texto = ancho_bloque - ancho_col_logo
                padding = 0.15 * self.PUNTOS_POR_CM

                tam_fuente = 7
                interlineado = tam_fuente + 1
                lineas = simpleSplit(
                    self.TEXTO_FOOTER,
                    "Helvetica",
                    tam_fuente,
                    ancho_col_texto - (2 * padding),
                )

                y_texto = y_base_footer + (len(lineas) - 1) * interlineado
                overlay.setFont("Helvetica", tam_fuente)
                for linea in lineas:
                    overlay.drawString(x_bloque + padding, y_texto, linea)
                    y_texto -= interlineado

                if imagen_logo is not None:
                    logo_w_original, logo_h_original = imagen_logo.getSize()
                    ancho_logo = 1.4 * 72
                    alto_logo = ancho_logo * (logo_h_original / logo_w_original)
                    x_logo = x_bloque + ancho_col_texto + (ancho_col_logo - ancho_logo) - padding
                    y_logo = y_base_footer + 0.05 * self.PUNTOS_POR_CM
                    overlay.drawImage(
                        imagen_logo,
                        x_logo,
                        y_logo,
                        width=ancho_logo,
                        height=alto_logo,
                        preserveAspectRatio=True,
                        mask="auto",
                    )

            overlay.save()
            overlay_stream.seek(0)

            pagina_overlay = PdfReader(overlay_stream).pages[0]
            pagina.merge_page(pagina_overlay)
            escritor.add_page(pagina)

        with open(ruta_pdf, "wb") as salida:
            escritor.write(salida)

    def _centrar_bloque_por_margenes(self):
        for seccion in self.doc.sections:
            margen_horizontal_total = seccion.left_margin + seccion.right_margin

            if self._ancho_extra_cm > 0:
                extra_emu = int(Cm(self._ancho_extra_cm))
                margen_horizontal_total = max(0, margen_horizontal_total - extra_emu)

            margen_izquierdo = int(margen_horizontal_total // 2)
            margen_derecho = int(margen_horizontal_total - margen_izquierdo)

            seccion.left_margin = margen_izquierdo
            seccion.right_margin = margen_derecho

    def _asegurar_margen_inferior_minimo(self):
        if self._margen_inferior_minimo_cm is None:
            return

        minimo = Cm(self._margen_inferior_minimo_cm)
        margen_superior_minimo = Cm(self._margen_superior_minimo_cm)

        for seccion in self.doc.sections:
            if seccion.bottom_margin >= minimo:
                continue

            delta = minimo - seccion.bottom_margin
            espacio_disponible_superior = seccion.top_margin - margen_superior_minimo

            if espacio_disponible_superior <= 0:
                seccion.bottom_margin = minimo
                continue

            transferencia = delta if delta <= espacio_disponible_superior else espacio_disponible_superior
            seccion.top_margin = seccion.top_margin - transferencia
            seccion.bottom_margin = seccion.bottom_margin + transferencia

            if seccion.bottom_margin < minimo:
                seccion.bottom_margin = minimo

    def _agregar_parrafo_formateado(self, contenido, bold=False):

        parrafo = self.doc.add_paragraph()
        parrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        parrafo.paragraph_format.line_spacing = 1.5

        if isinstance(contenido, dict):
            runs = contenido.get("runs", [])

            for texto, es_negrita in runs:
                run = parrafo.add_run(texto)
                run.font.name = "Calibri"
                run.font.size = Pt(9)
                run.bold = es_negrita
        elif isinstance(contenido, (list, tuple)):
            for texto, es_negrita in contenido:
                run = parrafo.add_run(texto)
                run.font.name = "Calibri"
                run.font.size = Pt(9)
                run.bold = es_negrita
        else:
            run = parrafo.add_run(contenido)
            run.font.name = "Calibri"
            run.font.size = Pt(9)
            run.bold = bold

        return parrafo

    def _marcar_parrafo_para_iguales(self, parrafo, caracter="=", indice=None):
        """Marca un párrafo para que, más adelante, se rellene con '=' hasta
        el margen derecho. No calcula nada aquí: el cálculo con una fuente
        aproximada (PIL) nunca coincide exactamente con cómo Word divide las
        líneas, y por eso quedaban huecos en el PDF final. El relleno real se
        hace en `_rellenar_iguales_en_documento_word`, con el documento ya
        abierto en Word, preguntándole a Word en qué línea queda el cursor
        después de cada inserción.
        """

        texto = parrafo.text.rstrip()
        if not texto:
            return

        if caracter == "=" and texto.endswith("="):
            return

        if indice is None:
            # Fallback robusto cuando no se pasa un indice explicito.
            indice = None
            for i, p in enumerate(self.doc.paragraphs):
                if p._element is parrafo._element:
                    indice = i
                    break

        if indice is None:
            return

        self._indices_pendientes_iguales.append((indice, caracter))

    def _rellenar_iguales_en_documento_word(self, documento):
        """Rellena con '=' el final de cada párrafo marcado, usando Word
        (ya abierto vía win32com) para medir con exactitud. Se insertan
        bloques de '=' cada vez más pequeños y, apenas Word reporta que el
        cursor saltó a una nueva línea, se deshace ese último bloque. Así el
        relleno queda exacto, porque lo mide el mismo motor que genera el
        PDF final (no una estimación con PIL/reportlab aparte).
        """

        if not self._indices_pendientes_iguales:
            return

        total_parrafos = documento.Paragraphs.Count

        for pendiente in self._indices_pendientes_iguales:
            if isinstance(pendiente, tuple):
                indice, caracter = pendiente
            else:
                indice, caracter = pendiente, "="

            numero_parrafo = indice + 1  # Los índices de Word son base 1.
            if numero_parrafo < 1 or numero_parrafo > total_parrafos:
                continue

            parrafo_word = documento.Paragraphs(numero_parrafo)
            rango = parrafo_word.Range
            rango.Collapse(self.WD_COLLAPSE_END)
            # Retrocede un caracter para quedar antes de la marca de fin de
            # párrafo (el "pilcrow"), y no después.
            rango.MoveEnd(self.WD_CHARACTER, -1)
            rango.Collapse(self.WD_COLLAPSE_END)

            lineas_base = parrafo_word.Range.ComputeStatistics(
                self.WD_STATISTIC_LINES
            )

            paso = 16
            while paso >= 1:
                while True:
                    rango.InsertAfter(caracter * paso)
                    rango.Collapse(self.WD_COLLAPSE_END)
                    lineas_actuales = parrafo_word.Range.ComputeStatistics(
                        self.WD_STATISTIC_LINES
                    )

                    if lineas_actuales != lineas_base:
                        # Este bloque hizo crecer el parrafo en una linea:
                        # se deshace y se prueba con un bloque más chico.
                        rango.MoveStart(self.WD_CHARACTER, -paso)
                        rango.Text = ""
                        break

                paso //= 2

    def _agregar_parrafo_con_lider(self, contenido, bold=False):
        """Agrega un párrafo y lo marca para que se rellene con '=' hasta el margen derecho."""

        parrafo = self.doc.add_paragraph()
        parrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        parrafo.paragraph_format.line_spacing = 1.5

        if isinstance(contenido, (list, tuple)):
            for texto, es_negrita in contenido:
                run = parrafo.add_run(texto)
                run.font.name = "Calibri"
                run.font.size = Pt(9)
                run.bold = es_negrita
        else:
            run = parrafo.add_run(contenido)
            run.font.name = "Calibri"
            run.font.size = Pt(9)
            run.bold = bold

        caracter_relleno = "="
        if isinstance(contenido, dict):
            caracter_relleno = contenido.get("fill_char", "=")

        self._marcar_parrafo_para_iguales(
            parrafo,
            caracter=caracter_relleno,
            indice=len(self.doc.paragraphs) - 1,
        )

        return parrafo

    def reemplazar_desde(
        self,
        texto_inicio,
        nuevo_texto,
        bold=False,
        con_lider_final=False,
        con_lider_indices=None,
    ):

        indices_con_lider = set(con_lider_indices or [])
        if isinstance(texto_inicio, str):
            textos_inicio = [texto_inicio]
        elif isinstance(texto_inicio, Iterable):
            textos_inicio = list(texto_inicio)
        else:
            textos_inicio = [texto_inicio]

        parrafos = list(self.doc.paragraphs)
        parrafo_inicio = None

        for p in parrafos:
            if any(texto in p.text for texto in textos_inicio):
                parrafo_inicio = p
                break

        if parrafo_inicio is None:
            if isinstance(nuevo_texto, list):
                for i, contenido in enumerate(nuevo_texto):
                    usar_lider = (i in indices_con_lider) or (
                        con_lider_final and i == len(nuevo_texto) - 1
                    )
                    if usar_lider:
                        self._agregar_parrafo_con_lider(contenido)
                    else:
                        self._agregar_parrafo_formateado(contenido)
            else:
                self._agregar_parrafo_formateado(nuevo_texto, bold=bold)
            return False

        elemento = parrafo_inicio._element
        cuerpo = elemento.getparent()

        while elemento is not None and not elemento.tag.endswith("}sectPr"):
            siguiente = elemento.getnext()
            cuerpo.remove(elemento)
            elemento = siguiente

        if isinstance(nuevo_texto, list):
            for i, contenido in enumerate(nuevo_texto):
                usar_lider = (i in indices_con_lider) or (
                    con_lider_final and i == len(nuevo_texto) - 1
                )
                if usar_lider:
                    self._agregar_parrafo_con_lider(contenido)
                else:
                    parrafo = self._agregar_parrafo_formateado(contenido)
                    caracter_relleno = "="
                    if isinstance(contenido, dict):
                        caracter_relleno = contenido.get("fill_char", "=")
                    self._marcar_parrafo_para_iguales(
                        parrafo,
                        caracter=caracter_relleno,
                        indice=len(self.doc.paragraphs) - 1,
                    )
        else:
            parrafo = self._agregar_parrafo_formateado(nuevo_texto, bold=bold)
            caracter_relleno = "="
            if isinstance(nuevo_texto, dict):
                caracter_relleno = nuevo_texto.get("fill_char", "=")
            self._marcar_parrafo_para_iguales(
                parrafo,
                caracter=caracter_relleno,
                indice=len(self.doc.paragraphs) - 1,
            )
        return True

    def aplicar_formato_global(self):
        # Centra el bloque usando margenes parejos sin cambiar el ancho util.
        self._centrar_bloque_por_margenes()
        self._asegurar_margen_inferior_minimo()

    def guardar(self, ruta):
        try:
            self.aplicar_formato_global()

            extension = os.path.splitext(ruta)[1].lower()

            # Se abre Word siempre (haya o no párrafos con "=" pendientes)
            # porque es la única forma de medir exactamente el mismo ancho
            # de línea que va a terminar en el documento final, sea .docx o
            # .pdf. Antes se guardaba el .docx directo con python-docx y
            # solo se pasaba por Word en el caso .pdf, por eso el relleno
            # con "=" (calculado aparte, con PIL) no coincidía con lo que
            # Word terminaba dibujando.
            with tempfile.TemporaryDirectory() as temp_dir:
                ruta_docx = os.path.join(temp_dir, "testimonio.docx")
                self.doc.save(ruta_docx)

                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False

                try:
                    documento = word.Documents.Open(os.path.abspath(ruta_docx))
                    try:
                        self._rellenar_iguales_en_documento_word(documento)

                        formato = (
                            self.FORMATO_PDF if extension == ".pdf" else self.FORMATO_DOCX
                        )
                        documento.SaveAs(os.path.abspath(ruta), FileFormat=formato)
                    finally:
                        documento.Close(False)
                finally:
                    word.Quit()

            if extension == ".pdf":
                self._estampar_firma_en_pdf(os.path.abspath(ruta))
        finally:
            if self.logo_temporal and os.path.exists(self.logo_temporal):
                os.remove(self.logo_temporal)
                self.logo_temporal = None