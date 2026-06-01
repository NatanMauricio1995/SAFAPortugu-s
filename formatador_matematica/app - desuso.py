from __future__ import annotations

import base64
import json
import os
import subprocess
import sys
import threading
import traceback
import webbrowser
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

BASE_DIR = Path(__file__).resolve().parent
LEGACY = None


def _load_legacy():
    global LEGACY
    if LEGACY is None:
        import legacy_backend as legacy  # type: ignore
        LEGACY = legacy
    return LEGACY


class BackendAPI:
    """API exposta ao HTML via pywebview."""

    def __init__(self):
        self.logs: list[str] = []

    def _log(self, msg: str):
        line = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
        self.logs.append(line)
        return line

    def ping(self):
        return {"ok": True, "message": "Backend Python conectado."}

    def get_logs(self):
        return {"ok": True, "logs": self.logs[-500:]}

    def choose_file(self, title="Selecionar arquivo", file_types=None):
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
            path = filedialog.askopenfilename(title=title, filetypes=file_types or [("Todos", "*.*")])
            root.destroy()
            return {"ok": True, "path": path or ""}
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}

    def choose_files(self, title="Selecionar arquivos", file_types=None):
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
            paths = filedialog.askopenfilenames(title=title, filetypes=file_types or [("Todos", "*.*")])
            root.destroy()
            return {"ok": True, "paths": list(paths)}
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}

    def choose_folder(self, title="Selecionar pasta"):
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
            path = filedialog.askdirectory(title=title)
            root.destroy()
            return {"ok": True, "path": path or ""}
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}

    def open_folder(self, path: str):
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

    def generate_txt(self, data: Dict[str, Any]):
        try:
            legacy = _load_legacy()
            data = {k: (v or "") for k, v in data.items()}
            required = ["item_base", "codigo_alvo"]
            missing = [k for k in required if not str(data.get(k, "")).strip()]
            if missing:
                return {"ok": False, "error": "Campos obrigatórios ausentes: " + ", ".join(missing)}
            text = legacy.TXT_GENERATOR_TEMPLATE.format(**data)
            self._log("TXT gerado com sucesso.")
            return {"ok": True, "text": text}
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}

    def format_item_to_word(self, raw_text: str, output_path: str):
        try:
            legacy = _load_legacy()
            if not raw_text.strip():
                return {"ok": False, "error": "Cole o texto bruto do item."}
            if not output_path.strip():
                return {"ok": False, "error": "Informe o caminho de saída .docx."}
            out = Path(output_path)
            out.parent.mkdir(parents=True, exist_ok=True)
            data = legacy.itemfmt_parse_input_text(raw_text)
            legacy.itemfmt_build_doc(data, str(out))
            self._log(f"Word formatado gerado: {out}")
            return {"ok": True, "path": str(out)}
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}

    def convert_safa(self, input_files: List[str], output_dir: str, ensinart: bool = True, mode: str = "todos", selection: str = ""):
        try:
            legacy = _load_legacy()
            if not input_files:
                return {"ok": False, "error": "Nenhum arquivo Word informado."}
            out = Path(output_dir or (Path(input_files[0]).parent / "saida_safa"))
            out.mkdir(parents=True, exist_ok=True)
            results = []
            for file in input_files:
                p = Path(file)
                if not p.exists():
                    results.append({"file": str(p), "ok": False, "error": "Arquivo não encontrado"})
                    continue
                conv = legacy.SafaConverter()
                # Usa process_file conforme a assinatura real do legacy_backend
                generated = conv.process_file(
                    p,
                    out,
                    ensinart=ensinart,
                    logger=lambda msg: self._log(msg),
                    item_selection=selection,
                )
                results.append({"file": str(p), "ok": True, "result": str(generated)})
            self._log("Conversão Padrão SAFA finalizada.")
            return {"ok": True, "results": results, "output_dir": str(out)}
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}

    def word_to_png(self, word_files: List[str], output_dir: str, dpi: int = 200, crop: bool = True, threshold: int = 128, extra_margin_px: int = 0):
        try:
            legacy = _load_legacy()
            if not word_files:
                return {"ok": False, "error": "Nenhum arquivo Word informado."}
            out = Path(output_dir or (Path(word_files[0]).parent / "png_export"))
            out.mkdir(parents=True, exist_ok=True)
            def log_cb(msg):
                self._log(msg)
            def progress_cb(done, total, status):
                self._log(f"{done}/{total} - {status}")
            # Assinatura: (word_files, base_output_dir, dpi, crop_vertical, threshold, extra_margin_px, log_callback, progress_callback)
            total, success = legacy.process_documents(
                [Path(x) for x in word_files],
                out,
                int(dpi),
                bool(crop),
                int(threshold),
                int(extra_margin_px),
                log_cb,
                progress_cb,
            )
            self._log(f"Montagem de imagens concluída: {success}/{total}.")
            return {"ok": True, "total": total, "success": success, "output_dir": str(out)}
        except Exception:
            return {"ok": False, "error": traceback.format_exc()}


def main():
    api = BackendAPI()
    html = str(BASE_DIR / "layout.html")
    try:
        import webview  # pip install pywebview
        webview.create_window("Adaptativa - 2026", html, js_api=api, width=1320, height=820, min_size=(1100, 700))
        webview.start(debug=False)
    except Exception:
        print("pywebview não encontrado ou falhou. Instale com: py -m pip install pywebview")
        print(traceback.format_exc())
        webbrowser.open(Path(html).as_uri())


if __name__ == "__main__":
    main()
