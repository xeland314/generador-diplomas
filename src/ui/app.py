import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from src.generator.excel_handler import (
    generar_plantilla_excel,
    leer_excel,
    validar_columnas,
)
from src.generator.pptx_generator import extraer_etiquetas, generar_certificados


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

        group_plantilla = ttk.LabelFrame(
            frame_principal, text=" 1. Plantilla PowerPoint ", padding=10
        )
        group_plantilla.pack(fill=tk.X, pady=(0, 10))

        ttk.Entry(
            group_plantilla,
            textvariable=self.ruta_plantilla,
            state="readonly",
            width=60,
        ).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(
            group_plantilla,
            text="Cargar Plantilla",
            command=self._cargar_plantilla,
        ).pack(side=tk.LEFT)

        self.lbl_etiquetas = ttk.Label(
            frame_principal,
            text="Etiquetas detectadas: (cargue una plantilla)",
            foreground="gray",
        )
        self.lbl_etiquetas.pack(anchor=tk.W, pady=(0, 5))

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
            group_excel,
            text="Cargar Excel",
            command=self._cargar_excel,
        ).pack(side=tk.LEFT)

        self.lbl_validacion = ttk.Label(
            frame_principal, text="", foreground="gray"
        )
        self.lbl_validacion.pack(anchor=tk.W, pady=(0, 5))

        self.frame_tabla = ttk.Frame(frame_principal)
        self.frame_tabla.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.tree = ttk.Treeview(self.frame_tabla, show="headings", height=8)
        scroll_y = ttk.Scrollbar(
            self.frame_tabla, orient=tk.VERTICAL, command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

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
            messagebox.showinfo("Listo", f"Plantilla Excel creada en:\n{ruta}")
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

        es_valido, faltantes, sobrantes = validar_columnas(
            columnas, self.etiquetas_pptx
        )

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
                    text=f"Excel valido: {len(self.datos_excel)} registros.",
                    foreground="green",
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
            messagebox.showerror(
                "Error", f"No se pudieron generar los certificados:\n{e}"
            )
