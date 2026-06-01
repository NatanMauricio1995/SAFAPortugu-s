"""
app.py — ponto de entrada do Formatador SAFA LP.

Uso:
    py app.py
"""

from __future__ import annotations

import json
import sys
import webbrowser
from pathlib import Path
from typing import Any, Dict

BASE_DIR = Path(__file__).resolve().parent
LAYOUT_DIR = BASE_DIR / "layout_interface"
INDEX_HTML = LAYOUT_DIR / "index.html"
OUTPUT_DIR = BASE_DIR / "output"
LOGS_DIR = BASE_DIR / "logs"


def garantir_pastas() -> None:
    for pasta in [LAYOUT_DIR, OUTPUT_DIR, LOGS_DIR]:
        pasta.mkdir(parents=True, exist_ok=True)


class Api:
    def ping(self) -> Dict[str, Any]:
        return {
            "ok": True,
            "mensagem": "Backend do Formatador SAFA LP ativo.",
            "base_dir": str(BASE_DIR),
        }

    def carregar_matrizes(self) -> Dict[str, Any]:
        try:
            from core.matriz_loader import MatrizLoader

            loader = MatrizLoader()
            return {
                "ok": True,
                "dados": loader.data,
                "matrizes": loader.get_matrizes_ativas(),
            }
        except Exception as exc:
            return {
                "ok": False,
                "erro": str(exc),
                "dados": None,
                "matrizes": [],
            }

    def listar_codigos(self, matriz_id: str, etapa_id: str) -> Dict[str, Any]:
        try:
            from core.matriz_loader import MatrizLoader

            loader = MatrizLoader()
            return {
                "ok": True,
                "codigos": loader.get_codigos(matriz_id, etapa_id),
            }
        except Exception as exc:
            return {
                "ok": False,
                "erro": str(exc),
                "codigos": [],
            }

    def buscar_codigo(self, matriz_id: str, etapa_id: str, codigo_digitado: str) -> Dict[str, Any]:
        try:
            from core.matriz_loader import MatrizLoader

            loader = MatrizLoader()
            encontrado = loader.buscar_codigo(matriz_id, etapa_id, codigo_digitado)
            if not encontrado:
                return {
                    "ok": False,
                    "erro": "Código não encontrado para a matriz e etapa selecionadas.",
                    "codigo": None,
                }
            return {"ok": True, "codigo": encontrado}
        except Exception as exc:
            return {"ok": False, "erro": str(exc), "codigo": None}

    def gerar_prompt(self, parametros: Dict[str, Any]) -> Dict[str, Any]:
        try:
            from core.prompt_engine import PromptEngine

            prompt = PromptEngine().gerar_prompt(parametros)
            return {"ok": True, "prompt": prompt}
        except Exception as exc:
            return {"ok": False, "erro": str(exc), "prompt": ""}

    def formatar_item_docx(self, texto_bruto: str, caminho_saida: str = "") -> Dict[str, Any]:
        return {
            "ok": False,
            "erro": "Função ainda não conectada à interface. Implementar chamada de core/word_formatter.py.",
        }

    def processar_padrao_safa(self, parametros: Dict[str, Any]) -> Dict[str, Any]:
        return {
            "ok": False,
            "erro": "Função ainda não conectada à interface. Implementar chamada de core/safa_processor.py.",
        }

    def converter_docx_para_png(self, parametros: Dict[str, Any]) -> Dict[str, Any]:
        return {
            "ok": False,
            "erro": "Função ainda não conectada à interface. Implementar chamada de core/png_converter.py.",
        }


def abrir_com_pywebview() -> None:
    import webview  # type: ignore

    api = Api()
    webview.create_window(
        title="Formatador SAFA — Língua Portuguesa",
        url=str(INDEX_HTML),
        js_api=api,
        width=1280,
        height=850,
        min_size=(1100, 750),
    )
    webview.start(debug=True)


def abrir_no_navegador_como_fallback() -> None:
    webbrowser.open(INDEX_HTML.as_uri())
    print("pywebview não está instalado. Abrindo interface no navegador.")
    print("Para usar a integração completa, instale as dependências com: py -m pip install pywebview")


def main() -> None:
    garantir_pastas()

    if not INDEX_HTML.exists():
        print(f"Erro: interface não encontrada em {INDEX_HTML}")
        sys.exit(1)

    try:
        abrir_com_pywebview()
    except ModuleNotFoundError:
        abrir_no_navegador_como_fallback()
    except Exception as exc:
        print(f"Erro ao abrir o programa: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
