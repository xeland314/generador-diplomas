# Generador de Diplomas

Aplicacion de escritorio con interfaz grafica (tkinter) para generar certificados en lote a partir de una plantilla de PowerPoint y un archivo de datos Excel.

## Instalacion

Requiere Python >= 3.10 y [uv](https://docs.astral.sh/uv/).

```bash
uv sync
```

## Ejecucion

```bash
uv run python main.py
```

## Funcionalidades

### 1. Plantilla PowerPoint

- Carga un archivo `.pptx` que contiene etiquetas con formato `{{nombre}}`, `{{curso}}`, etc.
- Detecta automaticamente todas las etiquetas presentes en la plantilla.
- Las etiquetas son dinamicas: no hay campos hardcodeados, se adaptan a cualquier plantilla.

### 2. Archivo Excel

- **Cargar Excel**: Importa un archivo `.xlsx` y valida que sus columnas coincidan con las etiquetas de la plantilla.
  - Si falta alguna columna requerida, muestra un error y no permite continuar.
  - Si hay columnas extra en el Excel, las ignora y las reporta.
- **Generar Plantilla Excel**: Crea un archivo `.xlsx` vacio con los encabezados correctos para que el usuario solo deba completar los datos.

### 3. Registros (CRUD)

Una vez cargada la plantilla, se pueden gestionar registros directamente desde la interfaz sin necesidad de un archivo Excel:

- **Agregar**: Abre un formulario dinamico con campos generados desde las etiquetas de la plantilla.
- **Editar**: Selecciona un registro existente y modifica sus valores.
- **Eliminar**: Elimina un registro seleccionado previa confirmacion.

### 4. Generacion de Certificados

- Genera un archivo `.pptx` con un slide por cada registro.
- Reemplaza todas las etiquetas `{{...}}` con los datos correspondientes.
- Preserva imagenes y formato de la plantilla original.

## Estructura del proyecto

```
generador-diplomas/
├── main.py                        # Punto de entrada
├── src/
│   ├── generator/
│   │   ├── pptx_generator.py      # Logica de PowerPoint
│   │   └── excel_handler.py       # Logica de Excel
│   └── ui/
│       ├── app.py                 # Interfaz grafica principal
│       └── form.py                # Formulario dinamico de registros
├── templates/                     # Plantillas pptx de ejemplo
├── assets/                        # Imagenes de ejemplo
└── pyproject.toml
```

## Dependencias

- [python-pptx](https://python-pptx.readthedocs.io/) - Manejo de archivos PowerPoint
- [openpyxl](https://openpyxl.readthedocs.io/) - Manejo de archivos Excel
