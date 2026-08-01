import subprocess
import sys
from pathlib import Path
from tkinter import Tk, messagebox

from updater import ensure_app_up_to_date


def _base_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _mostrar_error(mensaje: str):
    root = Tk()
    root.withdraw()
    messagebox.showerror("Generador de Testimonios", mensaje)
    root.destroy()


def main():
    _base_dir()  # Mantiene compatibilidad si se necesita ruta del ejecutable.
    try:
        resultado = ensure_app_up_to_date()
    except Exception as exc:
        _mostrar_error(f"Error al buscar/instalar actualización: {exc}")
        return

    app_path = resultado.executable_path

    if app_path is None or not app_path.exists():
        _mostrar_error(resultado.status)
        return

    if "fall" in resultado.status.lower() or "no se pudo" in resultado.status.lower():
        root = Tk()
        root.withdraw()
        messagebox.showwarning("Generador de Testimonios", resultado.status)
        root.destroy()

    subprocess.Popen([str(app_path)], cwd=str(app_path.parent))


if __name__ == "__main__":
    main()
