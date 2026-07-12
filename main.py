import sys


def main():
    if len(sys.argv) > 1:
        from src.cli import main as cli_main
        sys.exit(cli_main())
    else:
        import tkinter as tk
        from src.ui.app import Aplicacion

        root = tk.Tk()
        Aplicacion(root)
        root.mainloop()


if __name__ == "__main__":
    main()
