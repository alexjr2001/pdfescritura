from datetime import date, timedelta

from docx import Document

from generador import Generador
from utils import validar_fechas_no_futuras


def test_permite_fechas_iguales_o_anteriores_a_hoy():
    hoy = date.today()

    assert validar_fechas_no_futuras(hoy, [hoy, hoy - timedelta(days=1)])


def test_rechaza_cualquier_fecha_futura():
    hoy = date.today()

    assert not validar_fechas_no_futuras(hoy + timedelta(days=1), [hoy])
    assert not validar_fechas_no_futuras(hoy, [hoy + timedelta(days=1)])


def test_reemplazar_desde_marca_parrafo_para_relleno_con_igual(tmp_path):
    ruta_docx = tmp_path / "plantilla.docx"
    doc = Document()
    doc.add_paragraph("texto base")
    doc.save(ruta_docx)

    generador = Generador(str(ruta_docx))
    generador.reemplazar_desde(
        "texto base",
        [[("texto nuevo", True)]],
        con_lider_final=True,
    )

    assert generador._indices_pendientes_iguales
    assert generador._indices_pendientes_iguales[-1][1] == "="
