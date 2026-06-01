import subprocess
import urllib.request
import tempfile
import sys
from pathlib import Path

PYTHON_VERSION = "3.12"
APP_FILE = "app.py"

PYTHON_URL = "https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe"

# pacote que instala pelo pip : módulo usado para verificar
REQUIRED_PACKAGES = {
    "pywebview": "webview",
    "pillow": "PIL",
    "python-docx": "docx",
    "pymupdf": "fitz",
    "pywin32": "win32com",
    "tkinterdnd2": "tkinterdnd2",
}

def run(cmd, show=False):
    if show:
        return subprocess.run(cmd)
    return subprocess.run(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )

def python_ok():
    try:
        return run(["py", f"-{PYTHON_VERSION}", "--version"]).returncode == 0
    except Exception:
        return False

def instalar_python():
    print(f"Python {PYTHON_VERSION} não encontrado. Instalando...")

    temp_dir = Path(tempfile.gettempdir()) / "instalador_python_312"
    temp_dir.mkdir(parents=True, exist_ok=True)

    instalador = temp_dir / "python-3.12.7-amd64.exe"

    if not instalador.exists():
        print("Baixando Python...")
        urllib.request.urlretrieve(PYTHON_URL, instalador)

    print("Instalando Python...")
    result = subprocess.run([
        str(instalador),
        "/quiet",
        "InstallAllUsers=0",
        "PrependPath=1",
        "Include_pip=1",
        "Include_launcher=1"
    ])

    if result.returncode != 0:
        raise RuntimeError("Não foi possível instalar o Python.")

    if not python_ok():
        raise RuntimeError(
            "Python foi instalado, mas o comando py -3.12 ainda não foi localizado. "
            "Reinicie o computador e tente novamente."
        )

def modulo_ok(modulo):
    try:
        return run(["py", f"-{PYTHON_VERSION}", "-c", f"import {modulo}"]).returncode == 0
    except Exception:
        return False

def instalar_dependencias_faltantes():
    faltando = []

    print("Verificando dependências...")

    for pacote, modulo in REQUIRED_PACKAGES.items():
        if modulo_ok(modulo):
            print(f"OK - {pacote}")
        else:
            print(f"FALTANDO - {pacote}")
            faltando.append(pacote)

    if not faltando:
        print("Todas as dependências estão instaladas.")
        return

    print("Atualizando pip...")
    run(["py", f"-{PYTHON_VERSION}", "-m", "pip", "install", "--upgrade", "pip"], show=True)

    for pacote in faltando:
        print(f"Instalando {pacote}...")
        result = run(["py", f"-{PYTHON_VERSION}", "-m", "pip", "install", pacote], show=True)
        if result.returncode != 0:
            raise RuntimeError(f"Falha ao instalar: {pacote}")

def pasta_do_exe():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def abrir_app():
    base = pasta_do_exe()
    app = base / APP_FILE

    if not app.exists():
        raise FileNotFoundError(
            f"Não encontrei {APP_FILE} na mesma pasta deste executável.\n"
            f"Pasta verificada: {base}"
        )

    print(f"Abrindo: py -{PYTHON_VERSION} {APP_FILE}")
    subprocess.Popen(
        ["py", f"-{PYTHON_VERSION}", APP_FILE],
        cwd=str(base)
    )

def main():
    try:
        print("===================================")
        print("VERIFICADOR DO SISTEMA")
        print("===================================")

        if not python_ok():
            instalar_python()
        else:
            print(f"OK - Python {PYTHON_VERSION}")

        instalar_dependencias_faltantes()
        abrir_app()

        print("Sistema aberto.")
    except Exception as e:
        print("\nERRO:")
        print(e)
        input("\nPressione ENTER para fechar...")

if __name__ == "__main__":
    main()
