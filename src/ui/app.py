import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from src.generator.excel_handler import (
    generar_plantilla_excel,
    leer_excel,
    validar_columnas,
)
from src.generator.pptx_generator import extraer_etiquetas, generar_certificados
from src.ui.form import FormularioRegistro


class Aplicacion:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Diplomas")
        self.root.geometry("800x650")
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

        group_registros = ttk.LabelFrame(
            frame_principal, text=" 3. Registros ", padding=10
        )
        group_registros.pack(fill=tk.X, pady=(0, 5))

        self.btn_agregar = ttk.Button(
            group_registros, text="+ Agregar", command=self._agregar_registro
        )
        self.btn_agregar.pack(side=tk.LEFT, padx=(0, 5))

        self.btn_editar = ttk.Button(
            group_registros,
            text="Editar",
            command=self._editar_registro,
            state=tk.DISABLED,
        )
        self.btn_editar.pack(side=tk.LEFT, padx=(0, 5))

        self.btn_eliminar = ttk.Button(
            group_registros,
            text="Eliminar",
            command=self._eliminar_registro,
            state=tk.DISABLED,
        )
        self.btn_eliminar.pack(side=tk.LEFT, padx=(0, 5))

        self.lbl_conteo = ttk.Label(
            group_registros, text="0 registros", foreground="gray"
        )
        self.lbl_conteo.pack(side=tk.RIGHT)

        self.frame_tabla = ttk.Frame(frame_principal)
        self.frame_tabla.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.tree = ttk.Treeview(
            self.frame_tabla, show="headings", height=8, selectmode="browse"
        )
        scroll_y = ttk.Scrollbar(
            self.frame_tabla, orient=tk.VERTICAL, command=self.tree.yview
        )
        self.tree.configure(yscrollcommand=scroll_y.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<<TreeviewSelect>>", self._on_seleccion)

        btn_frame = ttk.Frame(frame_principal)
        btn_frame.pack(fill=tk.X)

        self.btn_generar = ttk.Button(
            btn_frame,
            text="Generar Certificados",
            command=self._generar_certificados,
            state=tk.DISABLED,
        )
        self.btn_generar.pack(side=tk.LEFT)

        self.lbl_estado = ttk.Label(
            btn_frame, text="", foreground="green"
        )
        self.lbl_estado.pack(side=tk.LEFT, padx=(15, 0))

        self.progress_frame = ttk.Frame(frame_principal)
        self.progress_frame.pack(fill=tk.X, pady=(8, 0))

        self.progress_bar = ttk.Progressbar(
            self.progress_frame, mode="determinate"
        )
        self.progress_bar.pack(fill=tk.X, side=tk.LEFT, expand=True, padx=(0, 10))

        self.lbl_progreso = ttk.Label(
            self.progress_frame, text="", width=20
        )
        self.lbl_progreso.pack(side=tk.RIGHT)

        self.progress_frame.pack_forget()

    # --- Plantilla ---

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
        self._refrescar_estado()

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

    # --- Excel ---

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
            msg_ok += f"\nColumnas ignoradas: {', '.join(sobrantes)}"
        self.lbl_validacion.config(text=msg_ok, foreground="green")
        self.btn_generar.config(state=tk.NORMAL)

        self._refrescar_tabla()

    # --- CRUD Registros ---

    def _agregar_registro(self):
        if not self.etiquetas_pptx:
            messagebox.showwarning(
                "Aviso", "Primero debe cargar una plantilla PowerPoint."
            )
            return

        form = FormularioRegistro(self.root, self.etiquetas_pptx, titulo="Agregar registro")
        self.root.wait_window(form)

        if form.resultado is None:
            return

        if not self.columnas_excel:
            self.columnas_excel = list(self.etiquetas_pptx)

        self.datos_excel.append(form.resultado)
        self._refrescar_tabla()
        self._refrescar_estado()

    def _editar_registro(self):
        seleccion = self.tree.selection()
        if not seleccion:
            return

        idx = self.tree.index(seleccion[0])
        valores = self.datos_excel[idx]

        form = FormularioRegistro(
            self.root,
            self.etiquetas_pptx,
            titulo="Editar registro",
            valores=valores,
        )
        self.root.wait_window(form)

        if form.resultado is None:
            return

        self.datos_excel[idx] = form.resultado
        self._refrescar_tabla()

        # Restaurar seleccion en la misma fila
        hijos = self.tree.get_children()
        if idx < len(hijos):
            self.tree.selection_set(hijos[idx])
            self.tree.focus(hijos[idx])

    def _eliminar_registro(self):
        seleccion = self.tree.selection()
        if not seleccion:
            return

        idx = self.tree.index(seleccion[0])
        registro = self.datos_excel[idx]
        resumen = ", ".join(f"{k}={v}" for k, v in registro.items())

        confirmar = messagebox.askyesno(
            "Confirmar eliminacion",
            f"Eliminar este registro?\n\n{resumen}",
        )
        if not confirmar:
            return

        self.datos_excel.pop(idx)
        self._refrescar_tabla()
        self._refrescar_estado()

    # --- Tabla ---

    def _refrescar_tabla(self):
        self.tree.delete(*self.tree.get_children())

        columnas = self.columnas_excel or self.etiquetas_pptx
        if not columnas:
            return

        self.tree["columns"] = columnas
        for col in columnas:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor=tk.CENTER)

        for fila in self.datos_excel:
            valores = [fila.get(c, "") for c in columnas]
            self.tree.insert("", tk.END, values=valores)

    def _on_seleccion(self, _event):
        tiene_seleccion = bool(self.tree.selection())
        estado = tk.NORMAL if tiene_seleccion else tk.DISABLED
        self.btn_editar.config(state=estado)
        self.btn_eliminar.config(state=estado)

    def _refrescar_estado(self):
        n = len(self.datos_excel)
        self.lbl_conteo.config(text=f"{n} registros")
        if n > 0 and self.etiquetas_pptx:
            self.btn_generar.config(state=tk.NORMAL)
        else:
            self.btn_generar.config(state=tk.DISABLED)

    # --- Generar ---

    def _generar_certificados(self):
        if not self.datos_excel:
            messagebox.showwarning("Aviso", "No hay datos cargados.")
            return

        ruta_salida = filedialog.asksaveasfilename(
            title="Guardar certificados",
            defaultextension=".pptx",
            filetypes=[("PowerPoint", "*.pptx")],
            initialfile="Certificados.pptx",
        )
        if not ruta_salida:
            return

        self._bloquear_ui(True)
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = len(self.datos_excel)
        self.lbl_progreso.config(text=f"0 / {len(self.datos_excel)}")
        self.progress_frame.pack(fill=tk.X, pady=(8, 0))

        thread = threading.Thread(
            target=self._worker_generar,
            args=(self.datos_excel, self.ruta_plantilla.get(), ruta_salida),
            daemon=True,
        )
        thread.start()

    def _worker_generar(self, datos, plantilla, salida):
        try:
            cantidad = generar_certificados(
                datos, plantilla, salida,
                on_progress=lambda act, tot: self.root.after(
                    0, self._actualizar_progreso, act, tot
                ),
            )
            self.root.after(0, self._finalizar_generacion, cantidad, salida, None)
        except Exception as e:
            self.root.after(0, self._finalizar_generacion, 0, salida, e)

    def _actualizar_progreso(self, actual, total):
        self.progress_bar["value"] = actual
        self.lbl_progreso.config(text=f"{actual} / {total}")

    def _finalizar_generacion(self, cantidad, salida, error):
        self.progress_frame.pack_forget()
        self._bloquear_ui(False)

        if error:
            messagebox.showerror(
                "Error", f"No se pudieron generar los certificados:\n{error}"
            )
        else:
            messagebox.showinfo(
                "Listo",
                f"Se generaron {cantidad} certificados en:\n{salida}",
            )
            self.lbl_estado.config(text=f"Certificados generados: {cantidad}")

    def _bloquear_ui(self, bloquear):
        estado = tk.DISABLED if bloquear else tk.NORMAL
        self.btn_generar.config(state=estado)
        self.btn_agregar.config(state=estado)
        self.btn_editar.config(state=estado if bloquear else tk.NORMAL)
        self.btn_eliminar.config(state=estado if bloquear else tk.NORMAL)
