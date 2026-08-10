import os
import shutil
import tempfile
import sys
from copy import deepcopy
from io import BytesIO
from collections.abc import Iterable

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm
from docx.shared import Emu
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.lib.utils import ImageReader, simpleSplit
from reportlab.pdfgen import canvas
import win32com.client

from utils import registrar_error_log


class Generador:

    FORMATO_PDF = 17
    FORMATO_DOCX = 12
    PUNTOS_POR_CM = 28.3464567

    # Constantes de la API de Word (win32com) usadas para el relleno con "=".
    WD_COLLAPSE_END = 0
    WD_CHARACTER = 1
    WD_GO_TO_PAGE = 1
    WD_GO_TO_ABSOLUTE = 1
    WD_TEXT_ORIENTATION_HORIZONTAL = 1
    WD_HEADER_FOOTER_FIRST_PAGE = 2
    WD_HEADER_FOOTER_PRIMARY = 1
    WD_HEADER_FOOTER_EVEN_PAGES = 3
    WD_RELATIVE_HORIZONTAL_POSITION_PAGE = 1
    WD_RELATIVE_VERTICAL_POSITION_PAGE = 1
    WD_WRAP_NONE = 3
    WD_FIRST_CHARACTER_LINE_NUMBER = 10
    WD_STATISTIC_PAGES = 2
    WD_STATISTIC_LINES = 1
    WD_TABLE_ALIGNMENT_LEFT = 0
    WD_ALIGN_VERTICAL_CENTER = 1
    MSO_TEXT_EFFECT_1 = 1
    MSO_FALSE = 0
    MSO_TRUE = -1
    MSO_SEND_BEHIND_TEXT = 5
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
        self._titulo_encabezado = ""
        self._docx_shape_id = 1000
        # Tuplas (indice_parrafo, caracter) de los párrafos que deben
        # rellenarse hasta el margen derecho. El relleno real se hace más
        # adelante, con Word abierto, para usar su medición exacta.
        self._indices_pendientes_iguales = []

    def establecer_titulo_encabezado(self, titulo):
        self._titulo_encabezado = (titulo or "").strip()

    def _aplicar_marca_agua_en_word(self, documento):
        if not self._titulo_encabezado:
            return

        texto = self._titulo_encabezado.upper()

        if documento.Sections.Count < 1:
            return

        seccion = documento.Sections(1)
        seccion.PageSetup.DifferentFirstPageHeaderFooter = self.MSO_TRUE
        encabezado = seccion.Headers(self.WD_HEADER_FOOTER_FIRST_PAGE)

        # Evita duplicar la marca de agua al guardar varias veces.
        for i in range(encabezado.Shapes.Count, 0, -1):
            forma = encabezado.Shapes.Item(i)
            if forma.Name == "MarcaAguaTestimonio":
                forma.Delete()

        marca = encabezado.Shapes.AddTextbox(
            self.WD_TEXT_ORIENTATION_HORIZONTAL,
            0,
            0,
            180,
            34,
            encabezado.Range,
        )

        marca.Name = "MarcaAguaTestimonio"
        texto_marca = marca.TextFrame.TextRange
        texto_marca.Text = texto
        texto_marca.Font.Name = "Calibri"
        texto_marca.Font.Size = 21
        texto_marca.Font.Bold = self.MSO_TRUE
        texto_marca.Font.Color = 8421504  # Gris plomo (RGB 128,128,128)
        marca.TextFrame.MarginLeft = 0
        marca.TextFrame.MarginRight = 0
        marca.TextFrame.MarginTop = 0
        marca.TextFrame.MarginBottom = 0
        marca.Line.Visible = self.MSO_FALSE
        marca.Fill.Visible = self.MSO_FALSE
        marca.Rotation = 0.0
        marca.LockAspectRatio = self.MSO_TRUE
        marca.WrapFormat.Type = self.WD_WRAP_NONE
        marca.RelativeHorizontalPosition = self.WD_RELATIVE_HORIZONTAL_POSITION_PAGE
        marca.RelativeVerticalPosition = self.WD_RELATIVE_VERTICAL_POSITION_PAGE
        marca.Top = 18

        ancho_pagina = seccion.PageSetup.PageWidth
        separacion_derecha = 8
        marca.Left = ancho_pagina - marca.Width - separacion_derecha

        marca.ZOrder(self.MSO_SEND_BEHIND_TEXT)

    def _configurar_footer_en_rango_word(self, documento, seccion, footer, ruta_logo):
        rango_footer = footer.Range
        rango_footer.Text = ""

        ancho_disponible = (
            seccion.PageSetup.PageWidth
            - seccion.PageSetup.LeftMargin
            - seccion.PageSetup.RightMargin
        )

        # En win32com Word todas las medidas se expresan en puntos.
        ancho_logo = 2.9 * self.PUNTOS_POR_CM
        padding_izquierdo = 18.0

        tabla = documento.Tables.Add(rango_footer, 1, 2)
        tabla.Borders.Enable = False
        tabla.Rows.Alignment = self.WD_TABLE_ALIGNMENT_LEFT

        ancho_col_texto = max(100.0, ancho_disponible - ancho_logo)
        # Word rechaza anchos fuera del rango [0, 1584] puntos.
        ancho_col_texto = min(1584.0, ancho_col_texto)
        ancho_logo = min(1584.0, max(1.0, ancho_logo))

        tabla.Columns(1).Width = ancho_col_texto
        tabla.Columns(2).Width = ancho_logo

        celda_texto = tabla.Cell(1, 1)
        celda_logo = tabla.Cell(1, 2)

        parrafo_texto = celda_texto.Range.Paragraphs(1)
        parrafo_texto.Range.Text = self.TEXTO_FOOTER
        parrafo_texto.Range.Font.Name = "Calibri"
        parrafo_texto.Range.Font.Size = 7
        parrafo_texto.Alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        parrafo_texto.LeftIndent = padding_izquierdo
        parrafo_texto.SpaceAfter = 0
        parrafo_texto.SpaceBefore = 0
        parrafo_texto.LineSpacingRule = 0

        celda_texto.VerticalAlignment = self.WD_ALIGN_VERTICAL_CENTER
        celda_logo.VerticalAlignment = self.WD_ALIGN_VERTICAL_CENTER

        if ruta_logo and os.path.exists(ruta_logo):
            try:
                rango_logo = celda_logo.Range
                rango_logo.Text = ""
                imagen_logo = rango_logo.InlineShapes.AddPicture(
                    FileName=os.path.abspath(ruta_logo),
                    LinkToFile=False,
                    SaveWithDocument=True,
                )
                imagen_logo.LockAspectRatio = self.MSO_TRUE
                imagen_logo.Width = 2.6 * self.PUNTOS_POR_CM
            except Exception:
                pass

        parrafo_logo = celda_logo.Range.Paragraphs(1)
        parrafo_logo.Alignment = WD_ALIGN_PARAGRAPH.RIGHT
        parrafo_logo.SpaceAfter = 0
        parrafo_logo.SpaceBefore = 0

    def _aplicar_footer_legal_en_word(self, documento):
        if documento.Sections.Count < 1:
            return

        ruta_logo = self._obtener_logo_suavizado()

        for i in range(1, documento.Sections.Count + 1):
            seccion = documento.Sections(i)
            seccion.PageSetup.DifferentFirstPageHeaderFooter = self.MSO_TRUE

            footer_primario = seccion.Footers(self.WD_HEADER_FOOTER_PRIMARY)
            self._configurar_footer_en_rango_word(
                documento,
                seccion,
                footer_primario,
                ruta_logo,
            )

            footer_primera_pagina = seccion.Footers(self.WD_HEADER_FOOTER_FIRST_PAGE)
            self._configurar_footer_en_rango_word(
                documento,
                seccion,
                footer_primera_pagina,
                ruta_logo,
            )

    @classmethod
    def _eliminar_primera_pagina_en_blanco_word(cls, documento):
        """Elimina la primera página si está vacía (sin texto, tablas ni figuras)."""

        total_paginas = documento.ComputeStatistics(cls.WD_STATISTIC_PAGES)
        if total_paginas <= 1:
            return False

        inicio_p1 = documento.GoTo(
            What=cls.WD_GO_TO_PAGE,
            Which=cls.WD_GO_TO_ABSOLUTE,
            Count=1,
        ).Start
        inicio_p2 = documento.GoTo(
            What=cls.WD_GO_TO_PAGE,
            Which=cls.WD_GO_TO_ABSOLUTE,
            Count=2,
        ).Start

        if inicio_p2 <= inicio_p1:
            return False

        rango_p1 = documento.Range(Start=inicio_p1, End=inicio_p2)
        texto_visible = (
            rango_p1.Text
            .replace("\r", "")
            .replace("\x0c", "")
            .replace("\x07", "")
            .strip()
        )

        if texto_visible:
            return False

        if rango_p1.Tables.Count > 0 or rango_p1.InlineShapes.Count > 0:
            return False

        try:
            if rango_p1.ShapeRange.Count > 0:
                return False
        except Exception:
            # ShapeRange lanza excepción cuando no hay formas flotantes.
            pass

        rango_p1.Delete()
        return True

    @classmethod
    def preparar_fuente_sin_primera_pagina_blanca(cls, ruta):
        """Crea una copia temporal del DOCX sin primera página en blanco.

        Retorna una tupla:
        (ruta_a_usar, ruta_temporal_a_eliminar_o_none, pagina_eliminada).
        """

        ruta_absoluta = os.path.abspath(ruta)

        archivo_temp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        ruta_temporal = archivo_temp.name
        archivo_temp.close()

        try:
            shutil.copy2(ruta_absoluta, ruta_temporal)
        except OSError:
            if os.path.exists(ruta_temporal):
                os.remove(ruta_temporal)
            return ruta_absoluta, None, False

        word = None
        documento = None

        try:
            word = cls._crear_word_application()

            documento = word.Documents.Open(os.path.abspath(ruta_temporal), ReadOnly=False)
            pagina_eliminada = cls._eliminar_primera_pagina_en_blanco_word(documento)
            if pagina_eliminada:
                documento.Save()

            return ruta_temporal, ruta_temporal, pagina_eliminada
        except Exception:
            if os.path.exists(ruta_temporal):
                os.remove(ruta_temporal)
            return ruta_absoluta, None, False
        finally:
            if documento is not None:
                documento.Close(False)
            if word is not None:
                word.Quit()

    def _base_recursos(self):
        return getattr(sys, "_MEIPASS", os.path.dirname(__file__))

    @staticmethod
    def _crear_word_application():
        word = win32com.client.DispatchEx("Word.Application")
        try:
            word.DisplayAlerts = 0
        except Exception:
            pass
        return word

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

    def _siguiente_shape_id_docx(self):
        self._docx_shape_id += 1
        return self._docx_shape_id

    @staticmethod
    def _vaciar_story_docx(story):
        for child in list(story._element):
            if child.tag.endswith("}p") or child.tag.endswith("}tbl"):
                story._element.remove(child)

    def _configurar_footer_en_story_docx(self, story, seccion, ruta_logo):
        self._vaciar_story_docx(story)

        ancho_tabla = seccion.page_width - seccion.left_margin - seccion.right_margin
        tabla = story.add_table(rows=1, cols=2, width=ancho_tabla)
        tabla.autofit = False

        ancho_logo = Cm(2.9)
        ancho_texto = max(Emu(0), ancho_tabla - ancho_logo)

        celda_texto = tabla.cell(0, 0)
        celda_logo = tabla.cell(0, 1)
        celda_texto.width = ancho_texto
        celda_logo.width = ancho_logo
        celda_texto.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        celda_logo.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        parrafo_texto = celda_texto.paragraphs[0]
        parrafo_texto.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        parrafo_texto.paragraph_format.left_indent = Pt(18)
        parrafo_texto.paragraph_format.space_before = Pt(0)
        parrafo_texto.paragraph_format.space_after = Pt(0)
        run_texto = parrafo_texto.add_run(self.TEXTO_FOOTER)
        run_texto.font.name = "Calibri"
        run_texto.font.size = Pt(7)

        parrafo_logo = celda_logo.paragraphs[0]
        parrafo_logo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        parrafo_logo.paragraph_format.space_before = Pt(0)
        parrafo_logo.paragraph_format.space_after = Pt(0)
        if ruta_logo and os.path.exists(ruta_logo):
            run_logo = parrafo_logo.add_run()
            run_logo.add_picture(ruta_logo, width=Cm(2.6))

    def _aplicar_footer_legal_en_docx(self, documento_docx):
        ruta_logo = self._obtener_logo_suavizado()

        for seccion in documento_docx.sections:
            seccion.different_first_page_header_footer = True
            self._configurar_footer_en_story_docx(seccion.footer, seccion, ruta_logo)
            self._configurar_footer_en_story_docx(seccion.first_page_footer, seccion, ruta_logo)
            self._configurar_footer_en_story_docx(seccion.even_page_footer, seccion, ruta_logo)

    def _crear_anchor_para_inline_docx(self, inline, pos_x, pos_y, nombre):
        anchor = OxmlElement("wp:anchor")
        anchor.set("distT", "0")
        anchor.set("distB", "0")
        anchor.set("distL", "0")
        anchor.set("distR", "0")
        anchor.set("simplePos", "0")
        anchor.set("relativeHeight", "251658240")
        anchor.set("behindDoc", "1")
        anchor.set("locked", "0")
        anchor.set("layoutInCell", "1")
        anchor.set("allowOverlap", "1")

        simple_pos = OxmlElement("wp:simplePos")
        simple_pos.set("x", "0")
        simple_pos.set("y", "0")
        anchor.append(simple_pos)

        position_h = OxmlElement("wp:positionH")
        position_h.set("relativeFrom", "page")
        pos_offset_h = OxmlElement("wp:posOffset")
        pos_offset_h.text = str(int(pos_x))
        position_h.append(pos_offset_h)
        anchor.append(position_h)

        position_v = OxmlElement("wp:positionV")
        position_v.set("relativeFrom", "page")
        pos_offset_v = OxmlElement("wp:posOffset")
        pos_offset_v.text = str(int(pos_y))
        position_v.append(pos_offset_v)
        anchor.append(position_v)

        anchor.append(deepcopy(inline.extent))

        effect_extent = OxmlElement("wp:effectExtent")
        effect_extent.set("l", "0")
        effect_extent.set("t", "0")
        effect_extent.set("r", "0")
        effect_extent.set("b", "0")
        anchor.append(effect_extent)

        anchor.append(OxmlElement("wp:wrapNone"))

        doc_pr = deepcopy(inline.docPr)
        doc_pr.set("id", str(self._siguiente_shape_id_docx()))
        doc_pr.set("name", nombre)
        anchor.append(doc_pr)

        c_nv_graphic_frame_pr = inline.xpath("./wp:cNvGraphicFramePr")
        if c_nv_graphic_frame_pr:
            anchor.append(deepcopy(c_nv_graphic_frame_pr[0]))

        anchor.append(deepcopy(inline.graphic))
        return anchor

    def _agregar_imagen_flotante_en_story_docx(self, story, ruta_imagen, ancho, alto, pos_x, pos_y, nombre):
        parrafo = story.add_paragraph()
        parrafo.paragraph_format.space_before = Pt(0)
        parrafo.paragraph_format.space_after = Pt(0)
        run = parrafo.add_run()
        inline_shape = run.add_picture(ruta_imagen, width=Emu(int(ancho)), height=Emu(int(alto)))
        drawing = run._r.xpath("./w:drawing")[0]
        inline = drawing.xpath("./wp:inline")[0]
        anchor = self._crear_anchor_para_inline_docx(inline, pos_x, pos_y, nombre)
        drawing.remove(inline)
        drawing.append(anchor)
        return inline_shape

    def _aplicar_firma_lateral_en_docx(self, documento_docx):
        ruta_firma = self._obtener_ruta_firma()
        if not ruta_firma or not os.path.exists(ruta_firma):
            return

        ancho_firma = int(Cm(1.4))
        alto_firma = ancho_firma

        try:
            with Image.open(ruta_firma) as imagen_firma:
                ancho_original, alto_original = imagen_firma.size
                if ancho_original > 0:
                    alto_firma = int(ancho_firma * (alto_original / ancho_original))
        except Exception:
            pass

        pos_x = int(Cm(-0.05))

        for indice_seccion, seccion in enumerate(documento_docx.sections, start=1):
            seccion.different_first_page_header_footer = True
            pos_y = int((int(seccion.page_height) - alto_firma) / 2)

            self._agregar_imagen_flotante_en_story_docx(
                seccion.header,
                ruta_firma,
                ancho_firma,
                alto_firma,
                pos_x,
                pos_y,
                f"FirmaMargenHeader_{indice_seccion}_primary",
            )
            self._agregar_imagen_flotante_en_story_docx(
                seccion.first_page_header,
                ruta_firma,
                ancho_firma,
                alto_firma,
                pos_x,
                pos_y,
                f"FirmaMargenHeader_{indice_seccion}_first",
            )
            self._agregar_imagen_flotante_en_story_docx(
                seccion.even_page_header,
                ruta_firma,
                ancho_firma,
                alto_firma,
                pos_x,
                pos_y,
                f"FirmaMargenHeader_{indice_seccion}_even",
            )

    def _postprocesar_docx_sin_com(self, ruta_docx):
        documento_docx = Document(ruta_docx)
        self._aplicar_footer_legal_en_docx(documento_docx)
        self._aplicar_firma_lateral_en_docx(documento_docx)
        documento_docx.save(ruta_docx)

    def _aplicar_firma_marca_agua_en_word(self, documento):
        """Inserta la firma lateral izquierda en todas las páginas de Word."""

        ruta_firma = self._obtener_ruta_firma()
        if not ruta_firma or not os.path.exists(ruta_firma):
            return

        ancho_firma = 1.4 * self.PUNTOS_POR_CM
        alto_firma = ancho_firma

        try:
            with Image.open(ruta_firma) as imagen_firma:
                ancho_original, alto_original = imagen_firma.size
                if ancho_original > 0:
                    alto_firma = ancho_firma * (alto_original / ancho_original)
        except Exception:
            pass

        # Deja la firma en el margen izquierdo con un pequeño respiro.
        x_firma = -0.05 * self.PUNTOS_POR_CM
        if documento.Sections.Count < 1:
            return

        # Coordenada vertical global (idéntica para todo el documento).
        altura_pagina_base = documento.Sections(1).PageSetup.PageHeight
        y_firma_global = (altura_pagina_base - alto_firma) / 2.0

        # Limpia posibles firmas antiguas en cuerpo (estrategias previas).
        for j in range(documento.Shapes.Count, 0, -1):
            forma = documento.Shapes.Item(j)
            if str(forma.Name).startswith("FirmaMargenTestimonio"):
                forma.Delete()

        # Limpia firmas antiguas en headers para evitar "solo primera página".
        for indice_seccion in range(1, documento.Sections.Count + 1):
            seccion = documento.Sections(indice_seccion)
            for tipo_header in (
                self.WD_HEADER_FOOTER_PRIMARY,
                self.WD_HEADER_FOOTER_FIRST_PAGE,
                self.WD_HEADER_FOOTER_EVEN_PAGES,
            ):
                try:
                    header = seccion.Headers(tipo_header)
                    for j in range(header.Shapes.Count, 0, -1):
                        forma = header.Shapes.Item(j)
                        if str(forma.Name).startswith("FirmaMargenTestimonio"):
                            forma.Delete()
                except Exception:
                    pass

        def _insertar_en_pagina(pagina):
            ancla = documento.GoTo(
                What=self.WD_GO_TO_PAGE,
                Which=self.WD_GO_TO_ABSOLUTE,
                Count=pagina,
            )

            firma = documento.Shapes.AddPicture(
                os.path.abspath(ruta_firma),
                False,
                True,
                x_firma,
                y_firma_global,
                ancho_firma,
                alto_firma,
                ancla,
            )

            firma.Name = f"FirmaMargenTestimonio_{pagina}"
            firma.WrapFormat.Type = self.WD_WRAP_NONE
            firma.RelativeHorizontalPosition = self.WD_RELATIVE_HORIZONTAL_POSITION_PAGE
            firma.RelativeVerticalPosition = self.WD_RELATIVE_VERTICAL_POSITION_PAGE
            firma.LockAspectRatio = self.MSO_TRUE
            firma.Left = x_firma
            firma.Top = y_firma_global
            try:
                firma.LockAnchor = self.MSO_TRUE
            except Exception:
                pass
            try:
                firma.LayoutInCell = self.MSO_FALSE
            except Exception:
                pass
            try:
                firma.ZOrder(self.MSO_SEND_BEHIND_TEXT)
            except Exception:
                pass

        # Inserta y verifica cobertura de la primera a la última página.
        for _ in range(5):
            try:
                documento.Repaginate()
            except Exception:
                pass

            total_paginas = documento.ComputeStatistics(self.WD_STATISTIC_PAGES)
            paginas_con_firma = set()
            for j in range(1, documento.Shapes.Count + 1):
                nombre = str(documento.Shapes.Item(j).Name)
                if not nombre.startswith("FirmaMargenTestimonio_"):
                    continue
                sufijo = nombre.replace("FirmaMargenTestimonio_", "", 1)
                if sufijo.isdigit():
                    paginas_con_firma.add(int(sufijo))

            faltantes = [
                pagina
                for pagina in range(1, total_paginas + 1)
                if pagina not in paginas_con_firma
            ]
            if not faltantes:
                break

            for pagina in faltantes:
                _insertar_en_pagina(pagina)

        # Pase final defensivo para asegurar la ultima pagina.
        try:
            documento.Repaginate()
        except Exception:
            pass

        total_paginas_final = documento.ComputeStatistics(self.WD_STATISTIC_PAGES)
        ultima_tiene_firma = False
        nombre_ultima = f"FirmaMargenTestimonio_{total_paginas_final}"
        for j in range(1, documento.Shapes.Count + 1):
            if str(documento.Shapes.Item(j).Name) == nombre_ultima:
                ultima_tiene_firma = True
                break

        if not ultima_tiene_firma and total_paginas_final >= 1:
            _insertar_en_pagina(total_paginas_final)

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

    def guardar(self, ruta, avisar=None):
        def emitir_aviso(mensaje):
            registrar_error_log("Generador.guardar", mensaje, nivel="WARNING")
            if avisar is not None:
                avisar(mensaje)

        try:
            registrar_error_log(
                "Generador.guardar",
                f"Inicio de guardado. ruta={ruta!r}",
                nivel="INFO",
            )
            self.aplicar_formato_global()
            registrar_error_log("Generador.guardar", "Formato global aplicado.", nivel="INFO")

            extension = os.path.splitext(ruta)[1].lower()
            registrar_error_log(
                "Generador.guardar",
                f"Extensión detectada: {extension!r}",
                nivel="INFO",
            )

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
                registrar_error_log("Generador.guardar", f"DOCX temporal creado en {ruta_docx!r}.", nivel="INFO")
                ruta_docx_final_local = None
                ruta_docx_intermedio_local = None

                word = self._crear_word_application()
                registrar_error_log("Generador.guardar", "Instancia de Word abierta en segundo plano.", nivel="INFO")

                try:
                    documento = word.Documents.Open(os.path.abspath(ruta_docx))
                    registrar_error_log("Generador.guardar", "Documento temporal abierto en Word.", nivel="INFO")
                    try:
                        self._aplicar_marca_agua_en_word(documento)
                        registrar_error_log("Generador.guardar", "Marca de agua aplicada.", nivel="INFO")
                        self._rellenar_iguales_en_documento_word(documento)
                        registrar_error_log("Generador.guardar", "Relleno de '=' aplicado.", nivel="INFO")

                        formato = (
                            self.FORMATO_PDF if extension == ".pdf" else self.FORMATO_DOCX
                        )
                        if extension == ".docx":
                            archivo_temp_local = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
                            ruta_docx_intermedio_local = archivo_temp_local.name
                            archivo_temp_local.close()
                            documento.SaveAs(os.path.abspath(ruta_docx_intermedio_local), FileFormat=formato)
                        else:
                            documento.SaveAs(os.path.abspath(ruta), FileFormat=formato)
                        registrar_error_log(
                            "Generador.guardar",
                            f"Documento guardado en {os.path.abspath(ruta)!r} con formato {formato!r}.",
                            nivel="INFO",
                        )

                        if extension == ".docx":
                            emitir_aviso(
                                "El DOCX terminó de guardarse desde Word. "
                                "Si el archivo se ve incompleto, el problema quedó acotado al postprocesado COM."
                            )
                    finally:
                        documento.Close(False)
                finally:
                    word.Quit()
                    registrar_error_log("Generador.guardar", "Word cerrado.", nivel="INFO")

            if extension == ".docx":
                if not ruta_docx_intermedio_local or not os.path.exists(ruta_docx_intermedio_local):
                    raise FileNotFoundError(
                        "No se generó el DOCX intermedio local para el postprocesado."
                    )

                emitir_aviso(
                    "Word guardó el DOCX base correctamente. Ahora se aplicará footer y firma directamente sobre el archivo DOCX, sin una segunda pasada de Word COM."
                )
                postproceso_completo = True
                try:
                    registrar_error_log("Generador.guardar", "Iniciando postprocesado DOCX sin COM.", nivel="INFO")
                    self._postprocesar_docx_sin_com(ruta_docx_intermedio_local)
                    registrar_error_log("Generador.guardar", "Footer y firma insertados en el DOCX sin COM.", nivel="INFO")
                except Exception as exc:
                    postproceso_completo = False
                    registrar_error_log(
                        "Generador.guardar",
                        "Fallo el postprocesado DOCX sin COM.",
                        exc=exc,
                        nivel="ERROR",
                    )
                    emitir_aviso(
                        "No se pudo aplicar footer y firma directamente en el DOCX. Se conservará el DOCX base ya guardado. "
                        f"Error técnico: {exc!r}."
                    )

                if postproceso_completo:
                    try:
                        shutil.copy2(ruta_docx_intermedio_local, os.path.abspath(ruta))
                        registrar_error_log(
                            "Generador.guardar",
                            f"DOCX final copiado a {os.path.abspath(ruta)!r} desde la ruta local temporal.",
                            nivel="INFO",
                        )
                    except Exception as exc:
                        registrar_error_log(
                            "Generador.guardar",
                            "No se pudo copiar el DOCX final desde la ruta local temporal al destino.",
                            exc=exc,
                            nivel="ERROR",
                        )
                        emitir_aviso(
                            "El DOCX se procesó en local, pero no pudo copiarse a la ruta final. Se conservará el DOCX base ya guardado. "
                            f"Error técnico: {exc!r}."
                        )
                else:
                    registrar_error_log(
                        "Generador.guardar",
                        "Se omite la copia final del DOCX postprocesado porque Word COM falló; se conserva el DOCX base ya guardado.",
                        nivel="WARNING",
                    )

            if extension == ".pdf":
                registrar_error_log("Generador.guardar", "Iniciando estampado de firma en PDF.", nivel="INFO")
                self._estampar_firma_en_pdf(os.path.abspath(ruta))
                registrar_error_log("Generador.guardar", "Firma estampada en PDF.", nivel="INFO")
            registrar_error_log("Generador.guardar", "Guardado finalizado sin excepción.", nivel="INFO")
        except Exception as exc:
            registrar_error_log(
                "Generador.guardar",
                f"Error inesperado durante el guardado de {ruta!r}.",
                exc=exc,
                nivel="ERROR",
            )
            raise
        finally:
            if self.logo_temporal and os.path.exists(self.logo_temporal):
                os.remove(self.logo_temporal)
                self.logo_temporal = None
            if extension == ".docx" and ruta_docx_intermedio_local and os.path.exists(ruta_docx_intermedio_local):
                os.remove(ruta_docx_intermedio_local)