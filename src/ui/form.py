import tkinter as tk
from tkinter import ttk


class FormularioRegistro(tk.Toplevel):
    """Modal dinámico que genera campos según las etiquetas de la plantilla."""

    def __init__(self, parent, etiquetas, titulo="Registro", valores=None):
        super().__init__(parent)
        self.resultado = None
        self.campos = {}

        self.title(titulo)
        self.resizable(False, False)
        self.grab_set()

        frame = ttk.Frame(self, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            frame, text="Complete los campos:", font=("", 10, "bold")
        ).pack(anchor=tk.W, pady=(0, 10))

        for etiqueta in etiquetas:
            fila = ttk.Frame(frame)
            fila.pack(fill=tk.X, pady=3)

            ttk.Label(fila, text=f"{etiqueta}:", width=15, anchor=tk.W).pack(
                side=tk.LEFT
            )
            var = tk.StringVar(
                value=valores.get(etiqueta, "") if valores else ""
            )
            entry = ttk.Entry(fila, textvariable=var, width=40)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.campos[etiqueta] = var

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(15, 0))

        ttk.Button(btn_frame, text="Guardar", command=self._guardar).pack(
            side=tk.RIGHT, padx=(5, 0)
        )
        ttk.Button(btn_frame, text="Cancelar", command=self.destroy).pack(
            side=tk.RIGHT
        )

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
        self.geometry(f"+{x}+{y}")

    def _guardar(self):
        self.resultado = {
            nombre: var.get() for nombre, var in self.campos.items()
        }
        self.destroy()
