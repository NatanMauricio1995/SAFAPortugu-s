"""
main.py — Ponto de entrada da aplicação Adaptativa 2026
Expõe as funções do legacy_backend ao HTML via pywebview (window.pywebview.api).
"""

from __future__ import annotations

import os
import subprocess
import sys
import traceback
import webbrowser
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

BASE_DIR = Path(__file__).resolve().parent

# ── Carregamento lazy do legacy para não bloquear a UI ─────────────────────
_LEGACY = None


def _legacy():
    global _LEGACY
    if _LEGACY is None:
        import legacy_backend as lb  # type: ignore
        _LEGACY = lb
    return _LEGACY


# ══════════════════════════════════════════════════════════════════════════════
# API exposta ao HTML
# ══════════════════════════════════════════════════════════════════════════════

class BackendAPI:
    """Métodos chamados pelo HTML via window.pywebview.api.<método>()."""

    def __init__(self):
        self._logs: list[str] = []

    # ── Utilitários internos ───────────────────────────────────────────────

    def _log(self, msg: str) -> str:
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
        self._logs.append(line)
        return line

    # ── Ping / status ──────────────────────────────────────────────────────

    def ping(self) -> dict:
        """Verifica se o backend Python está ativo."""
        return {"ok": True, "message": "Backend Python conectado."}

    def get_logs(self) -> dict:
        """Retorna os últimos 500 logs acumulados."""
        return {"ok": True, "logs": self._logs[-500:]}

    # ── Diálogos de arquivo / pasta ────────────────────────────────────────

    def choose_file(self, title: str = "Selecionar arquivo", file_types=None) -> dict:
        """Abre diálogo para selecionar um único arquivo."""
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            path = filedialog.askopenfilename(
                title=title,
                filetypes=file_types or [("Todos os arquivos", "*.*")],
            )
            root.destroy()
            return {"ok": True, "path": path or ""}
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}

    def choose_files(self, title: str = "Selecionar arquivos", file_types=None) -> dict:
        """Abre diálogo para selecionar múltiplos arquivos."""
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            paths = filedialog.askopenfilenames(
                title=title,
                filetypes=file_types or [("Todos os arquivos", "*.*")],
            )
            root.destroy()
            return {"ok": True, "paths": list(paths)}
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}

    def choose_folder(self, title: str = "Selecionar pasta") -> dict:
        """Abre diálogo para selecionar uma pasta."""
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            path = filedialog.askdirectory(title=title)
            root.destroy()
            return {"ok": True, "path": path or ""}
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}

    def open_folder(self, path: str) -> dict:
        """Abre uma pasta no explorador de arquivos do sistema."""
        try:
            p = Path(path)
            if not p.exists():
                return {"ok": False, "error": "Pasta não encontrada."}
            if os.name == "nt":
                os.startfile(str(p))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(p)], check=False)
            else:
                subprocess.run(["xdg-open", str(p)], check=False)
            return {"ok": True}
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}

    # ── Gerador de TXT ─────────────────────────────────────────────────────

    def generate_txt(self, data: Dict[str, Any]) -> dict:
        """
        Gera o TXT estruturado usando o template do legacy_backend.
        Campos obrigatórios em `data`: item_base, codigo_alvo.
        """
        try:
            lb = _legacy()

            # Garante que todos os valores sejam strings (sem None)
            data = {k: str(v or "").strip() for k, v in data.items()}

            # Validação mínima
            missing = [k for k in ("item_base", "codigo_alvo") if not data.get(k)]
            if missing:
                return {
                    "ok": False,
                    "error": "Campos obrigatórios ausentes: " + ", ".join(missing),
                }

            # TXT_GENERATOR_TEMPLATE pode estar ausente em versões antigas do legacy
            template = getattr(lb, "TXT_GENERATOR_TEMPLATE", None)
            if template is None:
                return {
                    "ok": False,
                    "error": "TXT_GENERATOR_TEMPLATE não encontrado em legacy_backend.py.",
                }

            text = template.format(**data)
            self._log("TXT gerado com sucesso.")
            return {"ok": True, "text": text}
        except KeyError as e:
            return {"ok": False, "error": f"Chave ausente no template: {e}"}
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}

    # ── Formatador de item → Word ──────────────────────────────────────────

    def format_item_to_word(self, raw_text: str, output_path: str) -> dict:
        """
        Converte texto bruto no padrão SAFA para um arquivo .docx formatado.
        """
        try:
            lb = _legacy()

            if not raw_text.strip():
                return {"ok": False, "error": "Cole o texto bruto do item."}
            if not output_path.strip():
                return {"ok": False, "error": "Informe o caminho de saída .docx."}

            parse_fn = getattr(lb, "itemfmt_parse_input_text", None)
            build_fn = getattr(lb, "itemfmt_build_doc", None)
            if parse_fn is None or build_fn is None:
                return {
                    "ok": False,
                    "error": "Funções itemfmt_parse_input_text / itemfmt_build_doc "
                             "não encontradas em legacy_backend.py.",
                }

            out = Path(output_path)
            out.parent.mkdir(parents=True, exist_ok=True)

            data = parse_fn(raw_text)
            build_fn(data, str(out))

            self._log(f"Word formatado gerado: {out}")
            return {"ok": True, "path": str(out)}
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}

    # ── Padrão SAFA ────────────────────────────────────────────────────────

    def convert_safa(
        self,
        input_files: List[str],
        output_dir: str,
        ensinart: bool = True,
        selection: str = "",
    ) -> dict:
        """
        Processa um ou mais arquivos Word aplicando o Padrão SAFA.
        Cada arquivo gera uma subpasta em output_dir com os itens padronizados.
        """
        try:
            lb = _legacy()

            if not input_files:
                return {"ok": False, "error": "Nenhum arquivo Word informado."}

            base_out = Path(
                output_dir.strip() if output_dir and output_dir.strip()
                else str(Path(input_files[0]).parent / "saida_safa")
            )
            base_out.mkdir(parents=True, exist_ok=True)

            converter = lb.SafaConverter()
            results = []

            for file_str in input_files:
                p = Path(file_str)
                if not p.exists():
                    results.append({
                        "file": str(p),
                        "ok": False,
                        "error": "Arquivo não encontrado.",
                    })
                    continue

                file_out = base_out / p.stem
                report = converter.process_file(
                    p,
                    file_out,
                    ensinart=bool(ensinart),
                    logger=self._log,
                    item_selection=selection or "",
                )
                results.append({
                    "file": str(p),
                    "ok": True,
                    "report": report,
                })
                self._log(
                    f"{p.name} — {report['itens_ok']} itens gerados, "
                    f"{report['itens_com_erro']} com erro."
                )

            self._log("Conversão Padrão SAFA finalizada.")
            return {"ok": True, "results": results, "output_dir": str(base_out)}
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}

    # ── Word → PNG ─────────────────────────────────────────────────────────

    def word_to_png(
        self,
        word_files: List[str],
        output_dir: str,
        dpi: int = 200,
        crop: bool = True,
    ) -> dict:
        """
        Converte arquivos Word em imagens PNG via process_documents do legacy.
        Requer Microsoft Word instalado (Windows, usa COM/pywin32).
        """
        try:
            lb = _legacy()

            if not word_files:
                return {"ok": False, "error": "Nenhum arquivo Word informado."}

            out = Path(
                output_dir.strip() if output_dir and output_dir.strip()
                else str(Path(word_files[0]).parent / "png_export")
            )
            out.mkdir(parents=True, exist_ok=True)

            def _progress(done: int, total: int, status: str):
                self._log(f"[{done}/{total}] {status}")

            total, success = lb.process_documents(
                [Path(x) for x in word_files],
                out,
                int(dpi),
                bool(crop),
                _progress,
            )

            self._log(f"Montagem de imagens concluída: {success}/{total}.")
            return {
                "ok": True,
                "total": total,
                "success": success,
                "output_dir": str(out),
            }
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}


# ══════════════════════════════════════════════════════════════════════════════
# Inicialização da janela
# ══════════════════════════════════════════════════════════════════════════════

def main():
    api = BackendAPI()
    html_path = str(BASE_DIR / "layout.html")

    try:
        import webview  # pip install pywebview
        webview.create_window(
            "Adaptativa — 2026",
            html_path,
            js_api=api,
            width=1320,
            height=820,
            min_size=(1100, 700),
        )
        webview.start(debug=False)
    except ImportError:
        print("pywebview não encontrado. Instale com:")
        print("  py -m pip install pywebview")
        print("\nAbrindo o HTML no navegador (sem backend Python ativo)...")
        webbrowser.open(Path(html_path).as_uri())
    except Exception:
        print("Erro ao iniciar pywebview:")
        print(traceback.format_exc())
        webbrowser.open(Path(html_path).as_uri())


if __name__ == "__main__":
    main()
