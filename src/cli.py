import argparse
import sys

from src.generator.excel_handler import (
    generar_plantilla_excel,
    leer_excel,
    validar_columnas,
)
from src.generator.pptx_generator import extraer_etiquetas, generar_certificados


def cmd_generate(args):
    """Genera certificados PPTX a partir de una plantilla y un Excel."""
    etiquetas = extraer_etiquetas(args.plantilla)
    if not etiquetas:
        print("Error: la plantilla no contiene etiquetas {{...}}.", file=sys.stderr)
        return 1

    columnas, datos = leer_excel(args.datos)
    ok, faltantes, sobrantes = validar_columnas(columnas, etiquetas)

    if not ok:
        print("Error: el Excel no coincide con la plantilla.", file=sys.stderr)
        if faltantes:
            print(f"  Faltan columnas: {', '.join(faltantes)}", file=sys.stderr)
        if sobrantes:
            print(f"  Columnas ignoradas: {', '.join(sobrantes)}", file=sys.stderr)
        return 1

    n = generar_certificados(datos, args.plantilla, args.salida)
    print(f"Se generaron {n} certificados en '{args.salida}'.")
    return 0


def cmd_template(args):
    """Genera un Excel vacio con los encabezados de la plantilla."""
    etiquetas = extraer_etiquetas(args.plantilla)
    if not etiquetas:
        print("Error: la plantilla no contiene etiquetas {{...}}.", file=sys.stderr)
        return 1

    generar_plantilla_excel(etiquetas, args.salida)
    print(f"Plantilla Excel creada en '{args.salida}' con columnas: {', '.join(etiquetas)}")
    return 0


def cmd_inspect(args):
    """Muestra las etiquetas detectadas en una plantilla."""
    etiquetas = extraer_etiquetas(args.plantilla)
    if not etiquetas:
        print("No se encontraron etiquetas {{...}} en la plantilla.")
        return 0

    print(f"Etiquetas detectadas ({len(etiquetas)}):")
    for e in etiquetas:
        print(f"  - {e}")
    return 0


def build_parser():
    parser = argparse.ArgumentParser(
        prog="generador-diplomas",
        description="Generador de certificados a partir de plantillas PowerPoint.",
    )
    sub = parser.add_subparsers(dest="comando", required=True)

    # generate
    p_gen = sub.add_parser("generate", help="Generar certificados PPTX")
    p_gen.add_argument("-p", "--plantilla", required=True, help="Ruta al archivo .pptx")
    p_gen.add_argument("-d", "--datos", required=True, help="Ruta al archivo .xlsx")
    p_gen.add_argument("-o", "--salida", default="Certificados.pptx", help="Archivo de salida (default: Certificados.pptx)")
    p_gen.set_defaults(func=cmd_generate)

    # template
    p_tpl = sub.add_parser("template", help="Generar Excel plantilla vacio")
    p_tpl.add_argument("-p", "--plantilla", required=True, help="Ruta al archivo .pptx")
    p_tpl.add_argument("-o", "--salida", default="plantilla_datos.xlsx", help="Archivo de salida (default: plantilla_datos.xlsx)")
    p_tpl.set_defaults(func=cmd_template)

    # inspect
    p_ins = sub.add_parser("inspect", help="Mostrar etiquetas de una plantilla")
    p_ins.add_argument("-p", "--plantilla", required=True, help="Ruta al archivo .pptx")
    p_ins.set_defaults(func=cmd_inspect)

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()
    return args.func(args)
