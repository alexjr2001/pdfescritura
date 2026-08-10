import os
import sys
import traceback
from datetime import date, datetime
import re
import unicodedata


DIAS = {
	1: "UNO",
	2: "DOS",
	3: "TRES",
	4: "CUATRO",
	5: "CINCO",
	6: "SEIS",
	7: "SIETE",
	8: "OCHO",
	9: "NUEVE",
	10: "DIEZ",
	11: "ONCE",
	12: "DOCE",
	13: "TRECE",
	14: "CATORCE",
	15: "QUINCE",
	16: "DIECISEIS",
	17: "DIECISIETE",
	18: "DIECIOCHO",
	19: "DIECINUEVE",
	20: "VEINTE",
	21: "VEINTIUNO",
	22: "VEINTIDOS",
	23: "VEINTITRES",
	24: "VEINTICUATRO",
	25: "VEINTICINCO",
	26: "VEINTISEIS",
	27: "VEINTISIETE",
	28: "VEINTIOCHO",
	29: "VEINTINUEVE",
	30: "TREINTA",
	31: "TREINTA Y UNO",
}

MESES = {
	1: "ENERO",
	2: "FEBRERO",
	3: "MARZO",
	4: "ABRIL",
	5: "MAYO",
	6: "JUNIO",
	7: "JULIO",
	8: "AGOSTO",
	9: "SEPTIEMBRE",
	10: "OCTUBRE",
	11: "NOVIEMBRE",
	12: "DICIEMBRE",
}

_DIAS_INV = {v: k for k, v in DIAS.items()}
_MESES_INV = {v: k for k, v in MESES.items()}
_DECENAS = {
	"CUARENTA": 40,
	"CINCUENTA": 50,
	"SESENTA": 60,
	"SETENTA": 70,
	"OCHENTA": 80,
	"NOVENTA": 90,
}
_UNIDADES = {
	"UNO": 1,
	"DOS": 2,
	"TRES": 3,
	"CUATRO": 4,
	"CINCO": 5,
	"SEIS": 6,
	"SIETE": 7,
	"OCHO": 8,
	"NUEVE": 9,
}


def _normalizar(texto: str) -> str:
	texto = texto.upper()
	texto = unicodedata.normalize("NFD", texto)
	texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
	texto = re.sub(r"[^A-Z0-9 ]+", " ", texto)
	texto = re.sub(r"\s+", " ", texto).strip()
	return texto


def _texto_a_numero(texto: str):
	texto = _normalizar(texto)

	if texto in _DIAS_INV:
		return _DIAS_INV[texto]

	if texto in _DECENAS:
		return _DECENAS[texto]

	m = re.match(
		r"^(CUARENTA|CINCUENTA|SESENTA|SETENTA|OCHENTA|NOVENTA) Y (UNO|DOS|TRES|CUATRO|CINCO|SEIS|SIETE|OCHO|NUEVE)$",
		texto,
	)
	if m:
		return _DECENAS[m.group(1)] + _UNIDADES[m.group(2)]

	return None


def _anio_notarial_a_numero(texto: str):
	texto = _normalizar(texto)

	if texto == "DOS MIL":
		return 2000

	if not texto.startswith("DOS MIL "):
		return None

	resto = texto.replace("DOS MIL ", "", 1).strip()
	valor = _texto_a_numero(resto)
	if valor is None:
		return None

	return 2000 + valor


def fecha_notarial_a_date(texto: str):
	texto_norm = _normalizar(texto)

	m = re.search(
		r"A LOS ([A-Z ]+?) DIAS DEL MES DE ([A-Z]+) DEL ANO ([A-Z ]+?)(?: ANTE MI|$)",
		texto_norm,
	)

	if not m:
		return None

	dia_txt, mes_txt, anio_txt = m.group(1), m.group(2), m.group(3)
	dia = _texto_a_numero(dia_txt)
	mes = _MESES_INV.get(_normalizar(mes_txt))
	anio = _anio_notarial_a_numero(anio_txt)

	if not (dia and mes and anio):
		return None

	try:
		return date(anio, mes, dia)
	except ValueError:
		return None


def anio_a_texto(anio: int) -> str:
	if 2000 <= anio <= 2099:
		if anio == 2000:
			return "DOS MIL"

		decenas = {
			1: "DIEZ",
			2: "VEINTE",
			3: "TREINTA",
			4: "CUARENTA",
			5: "CINCUENTA",
			6: "SESENTA",
			7: "SETENTA",
			8: "OCHENTA",
			9: "NOVENTA",
		}

		unidades = anio % 100
		if unidades <= 31:
			return f"DOS MIL {DIAS[unidades]}"

		d = unidades // 10
		u = unidades % 10

		if u == 0:
			return f"DOS MIL {decenas[d]}"

		return f"DOS MIL {decenas[d]} Y {DIAS[u]}"

	return str(anio)


def fecha_a_texto_notarial(fecha: date) -> str:
	dia = DIAS[fecha.day]
	mes = MESES[fecha.month]
	anio = anio_a_texto(fecha.year)
	return f"{dia} DE {mes} DEL {anio}"


def fecha_a_texto_interfaz(fecha: date) -> str:
	mes = MESES[fecha.month]
	return f"{fecha.day} DE {mes} DEL {fecha.year}"


def validar_fechas_no_futuras(fecha_notario: date, fechas_firma: list[date]) -> bool:
	"""Devuelve True cuando todas las fechas son iguales o anteriores al día actual."""
	hoy = date.today()
	if fecha_notario > hoy:
		return False

	return all(fecha <= hoy for fecha in fechas_firma)


def ruta_errors_log() -> str:
	if getattr(sys, "frozen", False):
		base_dir = os.path.dirname(sys.executable)
	else:
		base_dir = os.getcwd()

	return os.path.join(base_dir, "errors.log")


def registrar_error_log(contexto: str, mensaje: str, exc: Exception | None = None, nivel: str = "INFO") -> None:
	marca_tiempo = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
	lineas = [f"[{marca_tiempo}] [{nivel}] {contexto}", mensaje.strip()]

	if exc is not None:
		lineas.append("Traceback:")
		lineas.extend(traceback.format_exception(type(exc), exc, exc.__traceback__))

	lineas.append("")

	try:
		with open(ruta_errors_log(), "a", encoding="utf-8") as archivo:
			archivo.write("\n".join(lineas))
	except Exception:
		# El registro no debe romper la generación.
		pass
