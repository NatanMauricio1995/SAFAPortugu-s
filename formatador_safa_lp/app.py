"""
app.py — ponto de entrada provisório do Formatador SAFA LP.

Uso:
    python app.py

Este arquivo abre a interface localizada em layout_interface/index.html.
Ele também expõe uma API mínima para que o front-end possa começar a ser testado.
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
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"
LOGS_DIR = BASE_DIR / "logs"


def garantir_pastas() -> None:
    for pasta in [LAYOUT_DIR, DATA_DIR, TEMPLATES_DIR, OUTPUT_DIR, LOGS_DIR]:
        pasta.mkdir(parents=True, exist_ok=True)


def criar_index_minimo_se_nao_existir() -> None:
    if INDEX_HTML.exists():
        return

    INDEX_HTML.write_text(
        """<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <title>Formatador SAFA LP</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 40px; }
    .card { border: 1px solid #ddd; border-radius: 12px; padding: 24px; max-width: 900px; }
    code { background: #f5f5f5; padding: 2px 6px; border-radius: 4px; }
  </style>
</head>
<body>
  <div class="card">
    <h1>Formatador SAFA — Língua Portuguesa</h1>
    <p>A interface principal ainda não foi criada em <code>layout_interface/index.html</code>.</p>
    <p>Este é apenas um teste para confirmar que o <code>app.py</code> está abrindo.</p>
  </div>
</body>
</html>
""",
        encoding="utf-8",
    )


class Api:
    def ping(self) -> Dict[str, Any]:
        return {
            "ok": True,
            "mensagem": "Backend do Formatador SAFA LP ativo.",
            "base_dir": str(BASE_DIR),
        }

    def carregar_matrizes(self) -> Dict[str, Any]:
        caminho = DATA_DIR / "matrizes_portugues.json"
        if not caminho.exists():
            return {
                "ok": False,
                "erro": "Arquivo data/matrizes_portugues.json não encontrado.",
                "dados": None,
            }

        try:
            return {
                "ok": True,
                "dados": json.loads(caminho.read_text(encoding="utf-8")),
            }
        except Exception as exc:
            return {
                "ok": False,
                "erro": f"Erro ao ler matrizes_portugues.json: {exc}",
                "dados": None,
            }

    def gerar_prompt(self, parametros: Dict[str, Any]) -> Dict[str, Any]:
        modo = str(parametros.get("modo_trabalho", "")).strip().lower()

        if "revis" in modo:
            template_path = TEMPLATES_DIR / "prompt_portugues_revisao_item.txt"
        else:
            template_path = TEMPLATES_DIR / "prompt_portugues_novo_item.txt"

        if not template_path.exists():
            return {
                "ok": False,
                "erro": f"Template não encontrado: {template_path.name}",
                "prompt": "",
            }

        template = template_path.read_text(encoding="utf-8")

        prompt = template
        for chave, valor in parametros.items():
            prompt = prompt.replace("{{" + chave + "}}", str(valor or ""))

        return {
            "ok": True,
            "prompt": prompt,
        }

    def formatar_item_docx(self, texto_bruto: str, caminho_saida: str = "") -> Dict[str, Any]:
        return {
            "ok": False,
            "erro": "Função ainda não implementada. Implementar em core/word_formatter.py.",
        }

    def processar_padrao_safa(self, parametros: Dict[str, Any]) -> Dict[str, Any]:
        return {
            "ok": False,
            "erro": "Função ainda não implementada. Implementar em core/safa_processor.py.",
        }

    def converter_docx_para_png(self, parametros: Dict[str, Any]) -> Dict[str, Any]:
        return {
            "ok": False,
            "erro": "Função ainda não implementada. Implementar em core/png_converter.py.",
        }


def abrir_com_pywebview() -> None:
    import webview  # type: ignore

    api = Api()
    webview.create_window(
        title="Formatador SAFA — Língua Portuguesa",
        url=str(INDEX_HTML),
        js_api=api,
        width=1200,
        height=800,
        min_size=(1000, 700),
    )
    webview.start(debug=True)


def abrir_no_navegador_como_fallback() -> None:
    webbrowser.open(INDEX_HTML.as_uri())
    print("pywebview não está instalado. Abrindo interface no navegador.")
    print("Para usar a integração completa, instale as dependências com:")
    print("pip install -r requirements.txt")


def main() -> None:
    garantir_pastas()
    criar_index_minimo_se_nao_existir()

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
