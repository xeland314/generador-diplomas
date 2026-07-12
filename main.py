import copy
import io
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import openpyxl
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def extraer_etiquetas(ruta_plantilla):
    """Extrae todas las etiquetas {{...}} de la plantilla PPTX."""
    prs = Presentation(ruta_plantilla)
    etiquetas = set()
    patron = re.compile(r"\{\{(.+?)\}\}")

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        encontradas = patron.findall(run.text)
                        etiquetas.update(encontradas)
    return sorted(etiquetas)


def duplicar_slide(prs, slide):
    """Duplica una diapositiva copiando formas y reconstruyendo
    las relaciones de las imágenes para que no se pierdan."""
    layout = slide.slide_layout
    nueva = prs.slides.add_slide(layout)

    for shape in list(nueva.shapes):
        el = shape._element
        el.getparent().remove(el)

    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            image_stream = io.BytesIO(shape.image.blob)
            nueva.shapes.add_picture(
                image_stream, shape.left, shape.top, shape.width, shape.height
            )
        else:
            nuevo_el = copy.deepcopy(shape._element)
            nueva.shapes._spTree.insert_element_before(nuevo_el, "p:extLst")

    return nueva


def reemplazar_etiquetas(slide, datos):
    """Reemplaza todas las etiquetas {{...}} del slide con los datos."""
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    for clave, valor in datos.items():
                        run.text = run.text.replace(f"{{{{{clave}}}}}", str(valor))


def leer_excel(ruta_excel):
    """Lee el Excel y retorna una lista de diccionarios (una por fila)."""
    wb = openpyxl.load_workbook(ruta_excel, read_only=True)
    ws = wb.active
    filas = list(ws.iter_rows(values_only=True))
    wb.close()

    if len(filas) < 2:
        raise ValueError("El archivo Excel debe tener al menos una fila de encabezados y una fila de datos.")

    encabezados = [str(c).strip() for c in filas[0]]
    datos = []
    for i, fila in enumerate(filas[1:], start=2):
        registro = {}
        for j, valor in enumerate(fila):
            if j < len(encabezados):
                registro[encabezados[j]] = valor if valor is not None else ""
        datos.append(registro)
    return encabezados, datos


def validar_columnas(columnas_excel, etiquetas_pptx):
    """Valida que las columnas del Excel coincidan con las etiquetas del PPTX.
    Retorna (es_valido, faltantes, sobrantes)."""
    set_excel = set(columnas_excel)
    set_pptx = set(etiquetas_pptx)
    faltantes = sorted(set_pptx - set_excel)
    sobrantes = sorted(set_excel - set_pptx)
    es_valido = len(faltantes) == 0
    return es_valido, faltantes, sobrantes


def generar_plantilla_excel(etiquetas, ruta_salida):
    """Genera un Excel con los encabezados de las etiquetas detectadas."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Participantes"
    for col, etiqueta in enumerate(etiquetas, start=1):
        celda = ws.cell(row=1, column=col, value=etiqueta)
        celda.font = openpyxl.styles.Font(bold=True)
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 25
    wb.save(ruta_salida)
    wb.close()


def generar_certificados(lista_participantes, ruta_plantilla, archivo_salida):
    """Genera el PPTX con un slide por cada participante."""
    prs = Presentation(ruta_plantilla)
    plantilla = prs.slides[0]

    for persona in lista_participantes:
        slide = duplicar_slide(prs, plantilla)
        reemplazar_etiquetas(slide, persona)

    id_slide = prs.slides._sldIdLst[0]
    prs.slides._sldIdLst.remove(id_slide)

    prs.save(archivo_salida)
    return len(lista_participantes)


class Aplicacion:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Diplomas")
        self.root.geometry("750x600")
        self.root.resizable(False, False)

        self.ruta_plantilla = tk.StringVar()
        self.etiquetas_pptx = []
        self.datos_excel = []
        self.columnas_excel = []

        self._construir_ui()

    def _construir_ui(self):
        frame_principal = ttk.Frame(self.root, padding=15)
        frame_principal.pack(fill=tk.BOTH, expand=True)

        # --- Seccion Plantilla ---
        group_plantilla = ttk.LabelFrame(
            frame_principal, text=" 1. Plantilla PowerPoint ", padding=10
        )
        group_plantilla.pack(fill=tk.X, pady=(0, 10))

        ttk.Entry(
            group_plantilla, textvariable=self.ruta_plantilla, state="readonly", width=60
        ).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(
            group_plantilla, text="Cargar Plantilla", command=self._cargar_plantilla
        ).pack(side=tk.LEFT)

        self.lbl_etiquetas = ttk.Label(
            frame_principal, text="Etiquetas detectadas: (cargue una plantilla)", foreground="gray"
        )
        self.lbl_etiquetas.pack(anchor=tk.W, pady=(0, 5))

        # --- Seccion Excel ---
        group_excel = ttk.LabelFrame(
            frame_principal, text=" 2. Archivo Excel de datos ", padding=10
        )
        group_excel.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(
            group_excel,
            text="Generar Plantilla Excel",
            command=self._generar_plantilla_excel,
        ).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(
            group_excel, text="Cargar Excel", command=self._cargar_excel
        ).pack(side=tk.LEFT)

        self.lbl_validacion = ttk.Label(
            frame_principal, text="", foreground="gray"
        )
        self.lbl_validacion.pack(anchor=tk.W, pady=(0, 5))

        # --- Tabla preview ---
        self.frame_tabla = ttk.Frame(frame_principal)
        self.frame_tabla.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.tree = ttk.Treeview(self.frame_tabla, show="headings", height=8)
        scroll_y = ttk.Scrollbar(
            self.frame_tabla, orient=tk.VERTICAL, command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # --- Boton generar ---
        self.btn_generar = ttk.Button(
            frame_principal,
            text="Generar Certificados",
            command=self._generar_certificados,
            state=tk.DISABLED,
        )
        self.btn_generar.pack(pady=(5, 0))

        self.lbl_estado = ttk.Label(
            frame_principal, text="", foreground="green"
        )
        self.lbl_estado.pack(pady=(5, 0))

    # --- Acciones ---

    def _cargar_plantilla(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar plantilla PowerPoint",
            filetypes=[("PowerPoint", "*.pptx")],
        )
        if not ruta:
            return

        try:
            etiquetas = extraer_etiquetas(ruta)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer la plantilla:\n{e}")
            return

        if not etiquetas:
            messagebox.showwarning(
                "Aviso",
                "No se encontraron etiquetas con formato {{etiqueta}} en la plantilla.",
            )

        self.ruta_plantilla.set(ruta)
        self.etiquetas_pptx = etiquetas
        self.lbl_etiquetas.config(
            text=f"Etiquetas detectadas ({len(etiquetas)}): {', '.join(etiquetas)}",
            foreground="black",
        )
        self._validar_si_hay_excel()

    def _generar_plantilla_excel(self):
        if not self.etiquetas_pptx:
            messagebox.showwarning(
                "Aviso", "Primero debe cargar una plantilla PowerPoint."
            )
            return

        ruta = filedialog.asksaveasfilename(
            title="Guardar plantilla Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="plantilla_datos.xlsx",
        )
        if not ruta:
            return

        try:
            generar_plantilla_excel(self.etiquetas_pptx, ruta)
            messagebox.showinfo(
                "Listo",
                f"Plantilla Excel creada en:\n{ruta}",
            )
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear el Excel:\n{e}")

    def _cargar_excel(self):
        if not self.etiquetas_pptx:
            messagebox.showwarning(
                "Aviso", "Primero debe cargar una plantilla PowerPoint."
            )
            return

        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not ruta:
            return

        try:
            columnas, datos = leer_excel(ruta)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el Excel:\n{e}")
            return

        es_valido, faltantes, sobrantes = validar_columnas(columnas, self.etiquetas_pptx)

        if not es_valido:
            msg = "El Excel no coincide con la plantilla.\n\n"
            if faltantes:
                msg += f"Faltan estas columnas en el Excel:\n  - {'  - '.join(faltantes)}\n\n"
            if sobrantes:
                msg += f"Estas columnas del Excel no existen en la plantilla:\n  - {'  - '.join(sobrantes)}"
            messagebox.showerror("Error de validacion", msg)
            return

        self.columnas_excel = columnas
        self.datos_excel = datos
        msg_ok = f"Excel valido: {len(datos)} registros, columnas correctas."
        if sobrantes:
            msg_ok += f"\nColumnas ignoradas (no estan en la plantilla): {', '.join(sobrantes)}"
        self.lbl_validacion.config(text=msg_ok, foreground="green")
        self.btn_generar.config(state=tk.NORMAL)

        self._mostrar_preview(columnas, datos)

    def _mostrar_preview(self, columnas, datos):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = columnas
        for col in columnas:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor=tk.CENTER)
        for fila in datos[:50]:
            valores = [fila.get(c, "") for c in columnas]
            self.tree.insert("", tk.END, values=valores)

    def _validar_si_hay_excel(self):
        """Re-valida si ya habia un Excel cargado previamente."""
        if self.columnas_excel and self.datos_excel:
            es_valido, faltantes, _ = validar_columnas(
                self.columnas_excel, self.etiquetas_pptx
            )
            if not es_valido:
                self.lbl_validacion.config(
                    text=f"El Excel cargado ya no coincide. Faltan: {', '.join(faltantes)}",
                    foreground="red",
                )
                self.btn_generar.config(state=tk.DISABLED)
            else:
                self.lbl_validacion.config(
                    text=f"Excel valido: {len(self.datos_excel)} registros.", foreground="green"
                )
                self.btn_generar.config(state=tk.NORMAL)

    def _generar_certificados(self):
        if not self.datos_excel:
            messagebox.showwarning("Aviso", "No hay datos de Excel cargados.")
            return

        ruta_salida = filedialog.asksaveasfilename(
            title="Guardar certificados",
            defaultextension=".pptx",
            filetypes=[("PowerPoint", "*.pptx")],
            initialfile="Certificados.pptx",
        )
        if not ruta_salida:
            return

        try:
            cantidad = generar_certificados(
                self.datos_excel, self.ruta_plantilla.get(), ruta_salida
            )
            messagebox.showinfo(
                "Listo",
                f"Se generaron {cantidad} certificados en:\n{ruta_salida}",
            )
            self.lbl_estado.config(text=f"Certificados generados: {cantidad}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron generar los certificados:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = Aplicacion(root)
    root.mainloop()
