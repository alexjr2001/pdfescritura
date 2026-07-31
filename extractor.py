import re
from docx import Document
from utils import fecha_notarial_a_date


class Escritura:

    PATRONES_EMPRESA = (
        r"\bS\.?\s*A\.?\s*C\.?\b",
        r"\bS\.?\s*A\.?\s*A\.?\b",
        r"\bS\.?\s*A\.?\b",
        r"\bE\.?\s*I\.?\s*R\.?\s*L\.?\b",
        r"\bS\.?\s*R\.?\s*L\.?\b",
        r"\bSOCIEDAD\s+ANONIMA\b",
        r"\bSOCIEDAD\s+COMERCIAL\b",
    )

    def __init__(self, ruta):
        self.doc = Document(ruta)

    def texto(self):
        return "\n".join(p.text for p in self.doc.paragraphs)

    def extraer_comparecientes(self):

        texto = self.texto()

        try:
            inicio = texto.index("II.- COMPARECEN")
            fin = texto.index("III.- FE DE IDENTIFICACIÓN")
            bloque = texto[inicio:fin]
        except ValueError:
            # Si no se encuentran encabezados, buscar en todo el documento.
            bloque = texto

        patron = r"\d+\.-\s+([A-ZÁÉÍÓÚÑ ]+);"

        return re.findall(patron, bloque)

    def extraer_notario(self):

        texto = self.texto()

        patron = r"ANTE MÍ,\s*([A-ZÁÉÍÓÚÑ ]+),\s*NOTARIO"

        m = re.search(patron, texto)

        if m:
            return m.group(1)

        return ""

    def extraer_fecha_escritura(self):
        return fecha_notarial_a_date(self.texto())

    def extraer_numero_escritura(self):

        texto = self.texto()
        patron = r"ESCRITURA\s*:\s*(\d+)"
        m = re.search(patron, texto)

        if m:
            return m.group(1)

        return ""

    def extraer_numero_acta(self):

        texto = self.texto()

        patrones = [
            r"ACTA\s*[:#-]\s*(\d+)",
            r"ACTA\s+VEHICULAR\s*(?:N[°º\.]?\s*)?(\d+)",
        ]

        for patron in patrones:
            m = re.search(patron, texto, flags=re.IGNORECASE)
            if m:
                return m.group(1)

        return ""

    def _es_empresa(self, texto):
        limpio = texto.upper()
        return any(re.search(patron, limpio) for patron in self.PATRONES_EMPRESA)

    def _normalizar_bloque_personas(self, texto):
        texto = texto.upper()
        texto = re.sub(r"\s+", " ", texto)
        texto = texto.replace(" Y CÓNYUGE", "")
        texto = texto.replace(" Y CONYUGE", "")
        texto = texto.replace(" CON CÓNYUGE", "")
        texto = texto.replace(" CON CONYUGE", "")
        return texto.strip()

    def _extraer_nombres_persona_desde_texto(self, texto):
        if not texto:
            return []

        bloque = self._normalizar_bloque_personas(texto)

        if self._es_empresa(bloque):
            return []

        nombres = re.findall(r"[A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){1,5}", bloque)
        resultado = []

        for nombre in nombres:
            if self._es_empresa(nombre):
                continue

            palabras = nombre.split()
            if len(palabras) < 2:
                continue

            if all(p in {"CONYUGE", "CÓNYUGE", "Y", "E"} for p in palabras):
                continue

            resultado.append(nombre.strip())

        return resultado

    def _agregar_nombre_unico(self, personas, vistos, nombre):
        nombre = re.sub(r"^(?:DON|DOÑA)\s+", "", nombre.strip())
        clave = re.sub(r"\s+", " ", nombre).strip()
        if clave and clave not in vistos:
            vistos.add(clave)
            personas.append(clave)

    def _extraer_nombres_de_bloque_acta(self, bloque):
        personas = []
        vistos = set()

        patrones_contextuales = [
            r"(?:DON|DOÑA)\s+([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){2,5}),\s+QUIEN\b",
            r"([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){2,5}),\s+QUIEN\b",
            r"REPRESENTADA\s+POR\s+([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){2,5}),\s*PERUAN",
            r"\bY\s+POR\s+([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){2,5}),\s*PERUAN",
            r"CASAD[OA]\s+CON\s+([A-ZÁÉÍÓÚÑ]+(?:\s+[A-ZÁÉÍÓÚÑ]+){2,5})(?:\s*\(|\s*,)",
        ]

        for patron in patrones_contextuales:
            for nombre in re.findall(patron, bloque):
                nombre = re.sub(r"\s+", " ", nombre).strip()
                if self._es_empresa(nombre):
                    continue
                self._agregar_nombre_unico(personas, vistos, nombre)

        return personas

    def extraer_firmantes_acta(self):
        texto = self.texto()
        candidatos = []

        patrones_bloque = [
            r"COMO\s+VENDEDOR(?:ES)?\s*:\s*(.+?)(?=Y\s+DE\s+LA\s+OTRA\s+PARTE\s+POR\s+SU\s+PROPIO\s+DERECHO\s+COMO\s+COMPRADOR(?:ES)?\s*:)",
            r"COMO\s+COMPRADOR(?:ES)?\s*:\s*(.+?)(?=LOS\s+COMPARECIENTES\s+SON)",
        ]

        for patron in patrones_bloque:
            for bloque in re.findall(patron, texto, flags=re.IGNORECASE | re.DOTALL):
                candidatos.append(re.sub(r"\s+", " ", bloque).strip())

        if not candidatos:
            return self.extraer_comparecientes()

        personas = []
        vistos = set()

        for bloque in candidatos:
            for nombre in self._extraer_nombres_de_bloque_acta(bloque):
                self._agregar_nombre_unico(personas, vistos, nombre)

        if personas:
            return personas

        return self.extraer_comparecientes()