import tkinter as tk
import re
import os
from datetime import date
from tkinter import filedialog, messagebox

import customtkinter as ctk
from tkcalendar import DateEntry

from extractor import Escritura
from generador import Generador
from utils import fecha_a_texto_interfaz, fecha_a_texto_notarial, validar_fechas_no_futuras

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class App:

    def __init__(self):

        self.root = ctk.CTk()
        self.root.title("Generador de Testimonios")
        self.root.geometry("720x750")

        self.personas = []
        self.firmantes = []
        self.fecha_widgets = []
        self.fecha_texto_vars = []
        self.fecha_notario_texto_var = tk.StringVar()
        self.fecha_base_documento = date.today()
        self.tipo_documento_var = tk.StringVar(value="ESCRITURA")
        self.notario = ""
        self.numero_documento = ""
        self.boton_generar = None
        self.loader_frame = None
        self.loader_bar = None
        self.loader_label = None

        self.crear_interfaz()

        self.root.mainloop()

    def crear_interfaz(self):

        # ── Cabecera ──────────────────────────────────────────────
        ctk.CTkLabel(
            self.root,
            text="Generador de Testimonios",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(pady=(18, 4))

        ctk.CTkLabel(
            self.root,
            text="Documento Word (.docx)",
            font=ctk.CTkFont(size=12),
            text_color="gray70",
        ).pack()

        tipo_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        tipo_frame.pack(fill="x", padx=24, pady=(8, 2))

        ctk.CTkLabel(
            tipo_frame,
            text="Tipo de documento",
            font=ctk.CTkFont(size=12),
            text_color="gray70",
        ).pack(side="left")

        ctk.CTkOptionMenu(
            tipo_frame,
            values=["ESCRITURA", "ACTA VEHICULAR"],
            variable=self.tipo_documento_var,
            width=180,
        ).pack(side="left", padx=(10, 0))

        # ── Fila: ruta + botón ────────────────────────────────────
        frame = ctk.CTkFrame(self.root, fg_color="transparent")
        frame.pack(fill="x", padx=24, pady=(6, 0))

        self.ruta = tk.StringVar()

        ctk.CTkEntry(
            frame,
            textvariable=self.ruta,
            placeholder_text="Seleccione un archivo…",
            height=36,
        ).pack(side="left", fill="x", expand=True)

        ctk.CTkButton(
            frame,
            text="Examinar",
            width=100,
            height=36,
            command=self.abrir,
        ).pack(side="left", padx=(8, 0))

        # ── Separador ─────────────────────────────────────────────
        ctk.CTkFrame(self.root, height=2, fg_color="#3a3a3a").pack(
            fill="x", padx=24, pady=14
        )

        # ── Área scrollable de personas ───────────────────────────
        self.frame_personas = ctk.CTkScrollableFrame(
            self.root,
            label_text="",
            fg_color="transparent",
        )
        self.frame_personas.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        acciones_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        acciones_frame.pack(fill="x", padx=24, pady=(0, 12))

        ctk.CTkButton(
            acciones_frame,
            text="Agregar firmante",
            height=36,
            command=self.agregar_firmante,
        ).pack(side="left")

    def _set_generando(self, generando, texto="Generando documentos..."):
        if generando:
            if self.loader_frame is None:
                self.loader_frame = ctk.CTkFrame(self.root, corner_radius=10)
                self.loader_frame.pack(fill="x", padx=24, pady=(0, 12))

                self.loader_label = ctk.CTkLabel(
                    self.loader_frame,
                    text=texto,
                    font=ctk.CTkFont(size=12, weight="bold"),
                )
                self.loader_label.pack(anchor="w", padx=14, pady=(10, 4))

                self.loader_bar = ctk.CTkProgressBar(self.loader_frame, mode="determinate")
                self.loader_bar.pack(fill="x", padx=14, pady=(0, 12))
            else:
                self.loader_label.configure(text=texto)
                self.loader_frame.pack(fill="x", padx=24, pady=(0, 12))

            if self.loader_bar is not None:
                self.loader_bar.set(0.08)

            if self.boton_generar is not None:
                self.boton_generar.configure(state="disabled")

            # Forzar repintado para que el loader se vea antes de procesar.
            self.root.update_idletasks()
            self.root.update()
            return

        if self.loader_frame is not None:
            self.loader_frame.pack_forget()
        if self.boton_generar is not None:
            self.boton_generar.configure(state="normal")

    def _actualizar_loader(self, progreso, texto=None):
        if self.loader_bar is None:
            return

        if texto and self.loader_label is not None:
            self.loader_label.configure(text=texto)

        valor = max(0.0, min(1.0, float(progreso)))
        self.loader_bar.set(valor)
        self.root.update_idletasks()
        self.root.update()

    def abrir(self):

        archivo = filedialog.askopenfilename(
            filetypes=[("Word", "*.docx")]
        )

        if not archivo:
            return

        self.ruta.set(archivo)

        self.cargar_personas(archivo)

    def cargar_personas(self, archivo):
        doc = Escritura(archivo)

        if self.tipo_documento_var.get() == "ACTA VEHICULAR":
            self.personas = doc.extraer_firmantes_acta()
            self.numero_documento = doc.extraer_numero_acta()
        else:
            self.personas = doc.extraer_comparecientes()
            self.numero_documento = doc.extraer_numero_escritura()

        self.notario = doc.extraer_notario() or "BORIS ANTONIO VILCA GUTIÉRREZ"
        fecha_base = doc.extraer_fecha_escritura() or date.today()
        self.fecha_base_documento = fecha_base
        self.firmantes = [
            {"nombre": persona, "fecha": fecha_base}
            for persona in self.personas
        ]
        self._renderizar_firmantes()

    def _renderizar_firmantes(self):

        for widget in self.frame_personas.winfo_children():
            widget.destroy()

        self.fecha_widgets.clear()
        self.fecha_texto_vars.clear()
        self.personas = [firmante["nombre"] for firmante in self.firmantes]

        for indice, firmante in enumerate(self.firmantes):

            card = ctk.CTkFrame(self.frame_personas, corner_radius=10)
            card.pack(fill="x", padx=8, pady=(8, 0))

            cabecera = ctk.CTkFrame(card, fg_color="transparent")
            cabecera.pack(fill="x", padx=16, pady=(12, 2))

            ctk.CTkLabel(
                cabecera,
                text=firmante["nombre"],
                font=ctk.CTkFont(size=13, weight="bold"),
                anchor="w",
            ).pack(side="left", fill="x", expand=True)

            ctk.CTkButton(
                cabecera,
                text="Eliminar",
                width=84,
                height=28,
                fg_color="#7a2e2e",
                hover_color="#9a3a3a",
                command=lambda i=indice: self._eliminar_firmante(i),
            ).pack(side="right")

            ctk.CTkLabel(
                card,
                text="Fecha de firma",
                font=ctk.CTkFont(size=11),
                text_color="gray70",
                anchor="w",
            ).pack(anchor="w", padx=16)

            cal = DateEntry(
                card,
                width=28,
                date_pattern="dd/mm/yyyy",
                locale="es_PE",
                font=("Arial", 13),
            )
            cal.pack(anchor="w", padx=16, pady=(2, 0), ipady=6)
            cal.set_date(firmante["fecha"])

            texto_fecha_var = tk.StringVar()
            ctk.CTkLabel(
                card,
                textvariable=texto_fecha_var,
                font=ctk.CTkFont(size=11),
                text_color="#4fc3f7",
                anchor="w",
            ).pack(anchor="w", padx=16, pady=(2, 12))

            cal.bind(
                "<<DateEntrySelected>>",
                lambda event, var=texto_fecha_var, widget=cal, i=indice: self._actualizar_fecha_firmante(var, widget, i)
            )
            self._actualizar_fecha_firmante(texto_fecha_var, cal, indice)

            self.fecha_widgets.append(cal)
            self.fecha_texto_vars.append(texto_fecha_var)

        # ── Sección inferior: Notario + Fojas + Botón ─────────────
        bottom = ctk.CTkFrame(self.frame_personas, corner_radius=10)
        bottom.pack(fill="x", padx=8, pady=12)

        ctk.CTkLabel(
            bottom,
            text="Fecha del Notario",
            font=ctk.CTkFont(size=12),
            text_color="gray70",
            anchor="w",
        ).pack(anchor="w", padx=16, pady=(12, 0))

        self.fecha_notario = DateEntry(
            bottom,
            width=28,
            date_pattern="dd/mm/yyyy",
            locale="es_PE",
            font=("Arial", 13),
        )
        self.fecha_notario.pack(anchor="w", padx=16, pady=(2, 0), ipady=6)
        self.fecha_notario.set_date(self.fecha_base_documento)

        ctk.CTkLabel(
            bottom,
            textvariable=self.fecha_notario_texto_var,
            font=ctk.CTkFont(size=11),
            text_color="#4fc3f7",
            anchor="w",
        ).pack(anchor="w", padx=16, pady=(2, 8))

        self.fecha_notario.bind(
            "<<DateEntrySelected>>",
            lambda event: self._actualizar_fecha_texto(self.fecha_notario_texto_var, self.fecha_notario)
        )
        self._actualizar_fecha_texto(self.fecha_notario_texto_var, self.fecha_notario)

        # ── Fojas ─────────────────────────────────────────────────
        fojas_frame = ctk.CTkFrame(bottom, fg_color="transparent")
        fojas_frame.pack(fill="x", padx=16, pady=(4, 0))

        col1 = ctk.CTkFrame(fojas_frame, fg_color="transparent")
        col1.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkLabel(
            col1,
            text="Foja Inicial",
            font=ctk.CTkFont(size=12),
            text_color="gray70",
            anchor="w",
        ).pack(anchor="w")

        self.foja1 = ctk.CTkEntry(col1, placeholder_text="Ej: 438", height=34)
        self.foja1.pack(fill="x")

        col2 = ctk.CTkFrame(fojas_frame, fg_color="transparent")
        col2.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            col2,
            text="Foja Final",
            font=ctk.CTkFont(size=12),
            text_color="gray70",
            anchor="w",
        ).pack(anchor="w")

        self.foja2 = ctk.CTkEntry(col2, placeholder_text="Ej: 440V", height=34)
        self.foja2.pack(fill="x")

        # ── Botón de generación unificada ─────────────────────────
        botones_frame = ctk.CTkFrame(bottom, fg_color="transparent")
        botones_frame.pack(fill="x", padx=16, pady=(16, 16))

        self.boton_generar = ctk.CTkButton(
            botones_frame,
            text="Generar testimonio y parte",
            height=42,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self.generar_testimonio_y_parte,
        )
        self.boton_generar.pack(side="left", fill="both", expand=True)

    def _actualizar_fecha_firmante(self, var_texto, calendario, indice):
        self._actualizar_fecha_texto(var_texto, calendario)
        if 0 <= indice < len(self.firmantes):
            self.firmantes[indice]["fecha"] = calendario.get_date()

    def _eliminar_firmante(self, indice):
        if not (0 <= indice < len(self.firmantes)):
            return

        nombre = self.firmantes[indice]["nombre"]
        if not messagebox.askyesno("Eliminar firmante", f"¿Eliminar a {nombre}?"):
            return

        del self.firmantes[indice]
        self._renderizar_firmantes()

    def _pedir_nombre_firmante(self):
        dialogo = ctk.CTkToplevel(self.root)
        dialogo.title("Agregar firmante")
        dialogo.geometry("480x300")
        dialogo.minsize(480, 300)
        dialogo.resizable(False, False)
        dialogo.transient(self.root)
        dialogo.grab_set()

        resultado = {"nombre": None}

        contenedor = ctk.CTkFrame(dialogo, corner_radius=16)
        contenedor.pack(fill="both", expand=True, padx=18, pady=18)

        ctk.CTkLabel(
            contenedor,
            text="Agregar firmante",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(anchor="w", padx=18, pady=(14, 4))

        ctk.CTkLabel(
            contenedor,
            text="Escribe el nombre completo del firmante.",
            font=ctk.CTkFont(size=12),
            text_color="gray70",
        ).pack(anchor="w", padx=18, pady=(0, 10))

        entry = ctk.CTkEntry(
            contenedor,
            placeholder_text="Ej: Juan Pérez López",
            height=38,
        )
        entry.pack(fill="x", padx=18)

        mensaje_error = ctk.CTkLabel(
            contenedor,
            text="",
            font=ctk.CTkFont(size=11),
            text_color="#ff8a80",
        )
        mensaje_error.pack(anchor="w", padx=18, pady=(6, 0))

        botones = ctk.CTkFrame(contenedor, fg_color="transparent")
        botones.pack(fill="x", padx=18, pady=(14, 14))

        def cancelar():
            resultado["nombre"] = None
            dialogo.destroy()

        def aceptar():
            nombre = re.sub(r"\s+", " ", entry.get()).strip()
            if not nombre:
                mensaje_error.configure(text="Ingresa un nombre válido.")
                entry.focus_set()
                return

            resultado["nombre"] = nombre
            dialogo.destroy()

        ctk.CTkButton(
            botones,
            text="Cancelar",
            width=110,
            height=36,
            fg_color="#444444",
            hover_color="#555555",
            command=cancelar,
        ).pack(side="right", padx=(8, 0))

        ctk.CTkButton(
            botones,
            text="Agregar",
            width=120,
            height=36,
            command=aceptar,
        ).pack(side="right")

        dialogo.protocol("WM_DELETE_WINDOW", cancelar)
        dialogo.bind("<Return>", lambda event: aceptar())
        dialogo.bind("<Escape>", lambda event: cancelar())

        self.root.update_idletasks()
        dialogo.geometry(f"480x300+{self.root.winfo_rootx() + 120}+{self.root.winfo_rooty() + 100}")
        entry.focus_set()
        dialogo.lift()
        dialogo.attributes("-topmost", True)
        dialogo.after(150, lambda: dialogo.attributes("-topmost", False))
        self.root.wait_window(dialogo)

        return resultado["nombre"]

    def agregar_firmante(self):
        nombre = self._pedir_nombre_firmante()
        if not nombre:
            return

        nombre = re.sub(r"\s+", " ", nombre).strip()
        if not nombre:
            return

        fecha_base = self.fecha_notario.get_date() if hasattr(self, "fecha_notario") else date.today()
        self.firmantes.append({"nombre": nombre, "fecha": fecha_base})
        self._renderizar_firmantes()

    def _actualizar_fecha_texto(self, var_texto, calendario):

        fecha = calendario.get_date()
        var_texto.set(f"En texto: {fecha_a_texto_interfaz(fecha)}")

    def _crear_linea_asteriscos(self):
        # Se inserta un '*' inicial y Word completa exactamente hasta el final
        # de la linea durante el guardado (sin pasar a una segunda linea).
        return {
            "runs": [("*", False)],
            "fill_char": "*",
        }

    def _parsear_foja(self, valor):
        # Acepta formatos como: 123, 123V o 123 V.
        m = re.fullmatch(r"(\d+)\s*(V)?", valor.strip().upper())
        if not m:
            return None

        numero = int(m.group(1))
        es_vuelta = 1 if m.group(2) == "V" else 0
        return numero, es_vuelta

    def generar_testimonio_y_parte(self):

        self._set_generando(True)
        self._actualizar_loader(0.12, "Preparando documento fuente...")

        ruta_fuente, temporal_a_eliminar, pagina_eliminada = Generador.preparar_fuente_sin_primera_pagina_blanca(self.ruta.get())

        try:
            self._actualizar_loader(0.32, "Generando testimonio PDF...")
            ruta_testimonio = self.generar(mostrar_mensaje=False, ruta_fuente=ruta_fuente)
            if not ruta_testimonio:
                return

            self._actualizar_loader(0.68, "Generando parte Word/PDF...")
            rutas_parte = self.generar_parte(mostrar_mensaje=False, ruta_fuente=ruta_fuente)
            if not rutas_parte:
                return

            self._actualizar_loader(0.96, "Finalizando...")
        finally:
            if temporal_a_eliminar and os.path.exists(temporal_a_eliminar):
                os.remove(temporal_a_eliminar)
            self._set_generando(False)

        ruta_word_parte, ruta_pdf_parte = rutas_parte

        aviso_pagina = ""
        if pagina_eliminada:
            aviso_pagina = "\n\nAviso: se detectó y eliminó una primera página en blanco del Word fuente."

        messagebox.showinfo(
            "Éxito",
            "Documentos generados correctamente:\n"
            f"- Testimonio PDF: {ruta_testimonio}\n"
            f"- Parte Word: {ruta_word_parte}\n"
            f"- Parte PDF: {ruta_pdf_parte}"
            f"{aviso_pagina}"
        )

    def generar(self, mostrar_mensaje=True, ruta_fuente=None):

        if not self.ruta.get():

            messagebox.showerror(
                "Error",
                "Seleccione un documento."
            )

            return None

        if not self.firmantes:
            messagebox.showerror(
                "Error",
                "Agregue al menos un firmante."
            )
            return None

        ruta_origen = ruta_fuente
        temporal_a_eliminar = None
        if ruta_origen is None:
            ruta_origen, temporal_a_eliminar, _ = Generador.preparar_fuente_sin_primera_pagina_blanca(self.ruta.get())

        g = Generador(ruta_origen)
        g.establecer_titulo_encabezado("TESTIMONIO")

        texto_firmas = ""
        es_acta = self.tipo_documento_var.get() == "ACTA VEHICULAR"
        if es_acta:
            g.configurar_acta_vehicular()
        else:
            g.configurar_reserva_pie(3.2)

        fecha_notario_date = self.fecha_notario.get_date()
        fechas_firma = [cal.get_date() for cal in self.fecha_widgets]

        if any(fecha_notario_date < fecha_firma for fecha_firma in fechas_firma):
            messagebox.showerror(
                "Error",
                "La fecha del Notario debe ser igual o posterior a todas las fechas de firma."
            )
            return None

        if not validar_fechas_no_futuras(fecha_notario_date, fechas_firma):
            messagebox.showerror(
                "Error",
                "No se pueden generar documentos con fechas futuras. Corrija las fechas antes de continuar."
            )
            return None

        for firmante in self.firmantes:

            fecha = fecha_a_texto_notarial(firmante["fecha"])

            texto_firmas += (
                f"{firmante['nombre']} FIRMA EL {fecha}, UNA HUELLA DIGITAL. "
            )

        fecha_notario = fecha_a_texto_notarial(self.fecha_notario.get_date())
        fecha_hoy = fecha_a_texto_notarial(date.today())
        f1 = self.foja1.get().strip()
        f2 = self.foja2.get().strip()

        if f1 == "" or f2 == "":
            messagebox.showerror(
                "Error",
                "Ingrese fojas inicial y final."
            )
            return None

        foja_inicial = self._parsear_foja(f1)
        foja_final = self._parsear_foja(f2)

        if foja_inicial is None or foja_final is None:
            messagebox.showerror(
                "Error",
                "Formato de foja invalido. Use numero y opcional V (ejemplo: 438 o 438V)."
            )
            return None

        if foja_inicial > foja_final:
            messagebox.showerror(
                "Error",
                "La foja inicial debe ser menor o igual que la foja final."
            )
            return None

        if es_acta:
            texto = [
                self._crear_linea_asteriscos(),
                [
                    (
                        f"LA PRESENTE ACTA SE EXTIENDE DE LA FOJA SERIE NUMERO {f1} A {f2}. "
                        f"{texto_firmas}"
                        f"CONCLUIDO EL PROCESO DE FIRMAS, SUSCRIBO EL PRESENTE INSTRUMENTO EL DIA {fecha_notario}, {self.notario}, NOTARIO DE AREQUIPA. "
                        f"ES COPIA DE LA ACTA VEHICULAR QUE CORRE EN MI REGISTRO "
                        f"CON FECHA {fecha_hoy}, A FOJAS {f1}-{f2} "
                        "Y A SOLICITUD DE PARTE INTERESADA EXPIDO EL PRESENTE "
                        "TESTIMONIO NOTARIAL ELECTRONICO, DE ACUERDO A LEY.",
                        True,
                    )
                ],
            ]
        else:
            texto = [
                [
                    (
                        "FE DE CONTENIDO Y LECTURA:",
                        True,
                    ),
                    (
                        " INSTRUIDOS LOS OTORGANTES DEL CONTENIDO DEL PRESENTE "
                        "INSTRUMENTO POR LA LECTURA QUE LES HIZO EL NOTARIO, SE "
                        "RATIFICAN EN SU CONTENIDO, PROCEDIENDO A FIRMAR JUNTO "
                        "CONMIGO, DE LO QUE DOY FE.=============================="
                        "============================================================",
                        False,
                    )
                ],
                [
                    (
                        f"{texto_firmas}"
                        f"CONCLUIDO EL PROCESO DE FIRMAS, SUSCRIBO EL PRESENTE INSTRUMENTO EL DIA {fecha_notario}, {self.notario}, NOTARIO DE AREQUIPA. "
                        f"ES COPIA DE LA ESCRITURA PUBLICA QUE CORRE EN MI REGISTRO "
                        f"CON FECHA {fecha_hoy}, A FOJAS {f1}-{f2} "
                        "Y A SOLICITUD DE PARTE INTERESADA EXPIDO EL PRESENTE "
                        "TESTIMONIO NOTARIAL ELECTRONICO, DE ACUERDO A LEY.",
                        True,
                    )
                ],
            ]

        anclas_reemplazo = (
            ["LA PRESENTE ACTA SE EXTIENDE", "EL PROCESO DE FIRMAS CONCLUYO", "FE DE CONTENIDO Y LECTURA"]
            if es_acta else
            ["FE DE CONTENIDO Y LECTURA"]
        )

        g.reemplazar_desde(
            anclas_reemplazo,
            texto,
            con_lider_final=es_acta,
            con_lider_indices=[] if es_acta else [0, 1],
        )

        fecha_generacion = date.today().strftime("%Y-%m-%d")
        carpeta_base = r"C:\DOC BORIS VILCA\TESTIMONIOS"
        carpeta_destino = os.path.join(carpeta_base, fecha_generacion)
        numero_documento = self.numero_documento.strip() or "SIN_NUMERO"
        prefijo = "AV" if es_acta else "EP"
        nombre_pdf = f"{prefijo} {numero_documento}.pdf"

        try:
            try:
                os.makedirs(carpeta_destino, exist_ok=True)
                ruta_salida = os.path.join(carpeta_destino, nombre_pdf)
                g.guardar(ruta_salida)
                if mostrar_mensaje:
                    messagebox.showinfo(
                        "Éxito",
                        f"PDF generado correctamente en:\n{ruta_salida}"
                    )
            except OSError:
                ruta_salida = nombre_pdf
                g.guardar(ruta_salida)
                if mostrar_mensaje:
                    messagebox.showwarning(
                        "Aviso",
                        "No se pudo usar la ruta en Z:. Se guardó en la carpeta actual:\n"
                        f"{os.path.abspath(ruta_salida)}"
                    )
                ruta_salida = os.path.abspath(ruta_salida)
        finally:
            if temporal_a_eliminar and os.path.exists(temporal_a_eliminar):
                os.remove(temporal_a_eliminar)

        return ruta_salida

    def generar_parte(self, mostrar_mensaje=True, ruta_fuente=None):

        if not self.ruta.get():
            messagebox.showerror("Error", "Seleccione un documento.")
            return None

        if not self.firmantes:
            messagebox.showerror(
                "Error",
                "Agregue al menos un firmante."
            )
            return None

        ruta_origen = ruta_fuente
        temporal_a_eliminar = None
        if ruta_origen is None:
            ruta_origen, temporal_a_eliminar, _ = Generador.preparar_fuente_sin_primera_pagina_blanca(self.ruta.get())

        g = Generador(ruta_origen)
        es_acta = self.tipo_documento_var.get() == "ACTA VEHICULAR"
        if es_acta:
            g.configurar_acta_vehicular()
        else:
            g.configurar_reserva_pie(3.2)

        fecha_notario_date = self.fecha_notario.get_date()
        fechas_firma = [cal.get_date() for cal in self.fecha_widgets]

        if any(fecha_notario_date < fecha_firma for fecha_firma in fechas_firma):
            messagebox.showerror(
                "Error",
                "La fecha del Notario debe ser igual o posterior a todas las fechas de firma."
            )
            return None

        if not validar_fechas_no_futuras(fecha_notario_date, fechas_firma):
            messagebox.showerror(
                "Error",
                "No se pueden generar documentos con fechas futuras. Corrija las fechas antes de continuar."
            )
            return None

        f1 = self.foja1.get().strip()
        f2 = self.foja2.get().strip()

        if f1 == "" or f2 == "":
            messagebox.showerror("Error", "Ingrese fojas inicial y final.")
            return None

        foja_inicial = self._parsear_foja(f1)
        foja_final = self._parsear_foja(f2)

        if foja_inicial is None or foja_final is None:
            messagebox.showerror(
                "Error",
                "Formato de foja invalido. Use numero y opcional V (ejemplo: 438 o 438V)."
            )
            return None

        if foja_inicial > foja_final:
            messagebox.showerror(
                "Error",
                "La foja inicial debe ser menor o igual que la foja final."
            )
            return None

        # Construir texto para PARTE
        texto_fe_titulo = "FE DE CONTENIDO Y LECTURA:"
        texto_fe_base = (
            " INSTRUIDOS LOS OTORGANTES DEL CONTENIDO DEL PRESENTE "
            "INSTRUMENTO POR LA LECTURA QUE LES HIZO EL NOTARIO, SE "
            "RATIFICAN EN SU CONTENIDO, PROCEDIENDO A FIRMAR JUNTO "
            "CONMIGO, DE LO QUE DOY FE.=========================================================================================="
        )
        texto_fe_cuerpo = (
            texto_fe_base
        )

        # Párrafo de fojas
        texto_fojas_base = (
            f"LA PRESENTE {'ACTA' if es_acta else 'ESCRITURA'} SE INICIA EN LA FOJA SERIE B Nª {f1} "
            f"TERMINA EN LA FOJA SERIE {f2}. DOY FE."
        )
        texto_fojas = texto_fojas_base

        # Línea de asteriscos independiente, antes del párrafo principal.
        linea_asteriscos = self._crear_linea_asteriscos()
        partes_firmas = []

        for firmante in self.firmantes:
            fecha = fecha_a_texto_notarial(firmante["fecha"]).upper()
            persona_upper = firmante["nombre"].upper()
            partes_firmas.append(
                f"UNA FIRMA Y HUELLA DIGITAL DE {persona_upper}, A LOS {fecha}"
            )

        # Firma del notario
        fecha_notario = fecha_a_texto_notarial(self.fecha_notario.get_date()).upper()
        notario_upper = self.notario.upper()

        firmas_parte = "; ".join(partes_firmas)
        firmas_parte += (
            f"; CONCLUYENDO EL PROCESO DE FIRMAS ({len(self.firmantes)}), "
            f"A LOS {fecha_notario}, DOY FE; "
            f"ANTE MÍ, UNA FIRMA Y SELLO DE {notario_upper}, ABOGADO NOTARIO DE AREQUIPA."
        )

        # Párrafo final
        fecha_hoy = fecha_a_texto_notarial(date.today()).upper()
        texto_final = f"DOY FE QUE ESTA TRANSCRIPCIÓN CONCUERDA CON EL INSTRUMENTO PÚBLICO MATRIZ QUE OBRA EN MI REGISTRO, EL QUE SE ENCUENTRA SUSCRITO POR LOS OTORGANTES Y AUTORIZADA POR MÍ. SE EXPIDE ESTE PARTE NOTARIAL EN FORMATO DIGITAL, EN AREQUIPA A LOS {fecha_hoy}."

        if es_acta:
            texto = [
                linea_asteriscos,
                [
                    (
                        f"LA PRESENTE ACTA SE EXTIENDE DE LA FOJA SERIE NUMERO {f1} A {f2}. "
                        f"{firmas_parte} {texto_final}",
                        True,
                    ),
                ],
            ]
        else:
            texto = [
                [(texto_fe_titulo, True), (texto_fe_cuerpo, False)],
                [(texto_fojas, False)],
                linea_asteriscos,
                [(firmas_parte + " " + texto_final, True)],
            ]

        anclas_reemplazo = (
            ["LA PRESENTE ACTA SE EXTIENDE", "EL PROCESO DE FIRMAS CONCLUYO", "FE DE CONTENIDO Y LECTURA"]
            if es_acta else
            ["FE DE CONTENIDO Y LECTURA"]
        )

        g.reemplazar_desde(
            anclas_reemplazo,
            texto,
            con_lider_final=True,
            con_lider_indices=[] if es_acta else [1],
        )

        fecha_generacion = date.today().strftime("%Y-%m-%d")
        carpeta_base = r"C:\DOC BORIS VILCA\TESTIMONIOS"
        carpeta_destino = os.path.join(carpeta_base, fecha_generacion)
        numero_documento = self.numero_documento.strip() or "SIN_NUMERO"
        tipo_nombre = "Acta" if es_acta else "Parte"
        nombre_parte = f"{tipo_nombre} {numero_documento}"
        if es_acta:
            nombre_parte += " - PARTE"

        try:
            try:
                os.makedirs(carpeta_destino, exist_ok=True)

                # Guardar como Word
                ruta_word = os.path.join(carpeta_destino, f"{nombre_parte}.docx")
                g.guardar(ruta_word)

                # Guardar como PDF
                ruta_pdf = os.path.join(carpeta_destino, f"{nombre_parte}.pdf")
                g.guardar(ruta_pdf)
                if mostrar_mensaje:
                    messagebox.showinfo(
                        "Éxito",
                        f"PARTE generado correctamente en:\n{carpeta_destino}"
                    )
            except OSError:
                ruta_word = f"{nombre_parte}.docx"
                ruta_pdf = f"{nombre_parte}.pdf"
                g.guardar(ruta_word)
                g.guardar(ruta_pdf)
                ruta_word = os.path.abspath(ruta_word)
                ruta_pdf = os.path.abspath(ruta_pdf)
                if mostrar_mensaje:
                    messagebox.showwarning(
                        "Aviso",
                        f"Se guardó en la carpeta actual:\n{os.path.abspath(carpeta_destino)}"
                    )
        finally:
            if temporal_a_eliminar and os.path.exists(temporal_a_eliminar):
                os.remove(temporal_a_eliminar)

        return ruta_word, ruta_pdf