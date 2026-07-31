import subprocess
import sys
from pathlib import Path
from tkinter import Tk, messagebox

from updater import download_latest_app, get_latest_release, has_new_version


APP_EXE_NAME = "GeneradorTestimonios.exe"


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
    base_dir = _base_dir()
    app_path = base_dir / APP_EXE_NAME

    latest = get_latest_release()

    if latest is not None and has_new_version():
        if not download_latest_app(app_path):
            _mostrar_error(
                "No se pudo descargar la versión nueva desde GitHub.\n"
                "Se intentará abrir la versión local si existe."
            )

    if not app_path.exists():
        _mostrar_error(
            f"No se encontró {APP_EXE_NAME} en la carpeta de instalación.\n"
            "Vuelva a instalar o revise la descarga del Release."
        )
        return

    subprocess.Popen([str(app_path)], cwd=str(base_dir))


if __name__ == "__main__":
    main()
