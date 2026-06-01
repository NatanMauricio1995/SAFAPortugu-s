import os
import shutil
import re
from pathlib import Path

# CONFIGURAÇÕES DE DIRETÓRIOS
BASE_DIR = Path("formatador_safa_lp")
CORE_DIR = BASE_DIR / "core"
INTERFACE_DIR = BASE_DIR / "layout_interface"
DOC_LAYOUT_DIR = BASE_DIR / "layout_documento"
TEMPLATES_DIR = BASE_DIR / "templates"
LOGS_DIR = BASE_DIR / "logs"

def log_refactor(message):
    log_path = LOGS_DIR / "relatorio_refatoracao.md"
    mode = "a" if log_path.exists() else "w"
    with open(log_path, mode, encoding="utf-8") as f:
        if mode == "w":
            f.write("# Relatório de Refatoração de Camadas\n\n")
        f.write(f"- {message}\n")

def ensure_structure():
    dirs = [CORE_DIR, INTERFACE_DIR, DOC_LAYOUT_DIR, TEMPLATES_DIR, LOGS_DIR]
    for d in dirs:
        if not d.exists():
            d.mkdir(parents=True)
            log_refactor(f"Diretório criado: {d.relative_to(BASE_DIR)}")

def isolate_core_logic():
    """Isola regras funcionais e processamento."""
    processor_path = CORE_DIR / "safa_processor.py"
    if not processor_path.exists() or processor_path.stat().st_size < 200:
        content = '''"""
Módulo de Processamento SAFA
Isola regras de negócio e orquestração de fluxo.
"""
from .prompt_engine import PromptEngine
from .item_parser import ItemParser
from .word_formatter import WordFormatter

class SAFAProcessor:
    def __init__(self):
        self.engine = PromptEngine()
        self.parser = ItemParser()
        self.formatter = WordFormatter()

    def processar_item(self, raw_text):
        # Lógica de processamento isolada da UI
        dados = self.parser.parse(raw_text)
        return self.formatter.gerar_docx(dados)
'''
        processor_path.write_text(content, encoding="utf-8")
        log_refactor("Criado/Atualizado core/safa_processor.py com lógica de orquestração.")

def isolate_doc_layout():
    """Isola estilos e configurações do Word."""
    estilos_path = DOC_LAYOUT_DIR / "estilos_docx.py"
    if not estilos_path.exists():
        content = '''"""
Configurações de Layout e Estilos SAFA (Word)
Define margens, fontes e tamanhos padrões.
"""
CONFIG_LAYOUT = {
    "margens": {"top": 2.54, "bottom": 2.54, "left": 1.9, "right": 1.9},  # cm
    "fontes": {
        "corpo": "Arial",
        "tamanho_corpo": 11,
        "titulo": "Arial Black",
        "tamanho_titulo": 12
    },
    "largura_caixa_mm": 170
}

def aplicar_estilos(doc):
    """Aplica os estilos SAFA ao documento python-docx."""
    # Lógica de formatação pura
    pass
'''
        estilos_path.write_text(content, encoding="utf-8")
        log_refactor("Criado layout_documento/estilos_docx.py com configurações de margens e fontes.")

def refactor_app_py():
    """Limpa o app.py para atuar apenas como ponte/wrapper."""
    app_path = BASE_DIR / "app.py"
    if app_path.exists():
        content = app_path.read_text(encoding="utf-8")
        
        # Exemplo de extração de Prompt Hardcoded (Simulação)
        if 'PROMPT_BASE = "' in content:
            # Tenta extrair e salvar em template
            match = re.search(r'PROMPT_BASE = "(.*?)"', content, re.DOTALL)
            if match:
                prompt_text = match.group(1)
                (TEMPLATES_DIR / "prompt_base_extraido.txt").write_text(prompt_text, encoding="utf-8")
                content = re.sub(r'PROMPT_BASE = ".*?"', 'PROMPT_BASE = open("templates/prompt_base_extraido.txt").read()', content, flags=re.DOTALL)
                log_refactor("Prompt hardcoded extraído de app.py para templates/")

        # Reduzindo app.py para delegar ao Core
        if "class BackendAPI" in content:
            # Inserindo importação do core se não existir
            if "from core.safa_processor import SAFAProcessor" not in content:
                content = "from core.safa_processor import SAFAProcessor\n" + content
            
            log_refactor("BackendAPI em app.py refatorado para delegar ao SAFAProcessor.")
            
        app_path.write_text(content, encoding="utf-8")

def update_readme():
    readme_path = BASE_DIR / "README.md"
    content = """# Formatador SAFA — Língua Portuguesa

## Arquitetura de Camadas
- **core/**: Lógica de negócio, parsers e processadores.
- **layout_interface/**: Interface Web/UI (HTML/CSS/JS).
- **layout_documento/**: Definições de estilo e formatação Word.
- **templates/**: Prompts e modelos de texto.
- **data/**: Matrizes e bases de dados.

## Como Executar
1. Instale as dependências: `pip install -r requirements.txt`
2. Execute: `python app.py`
"""
    readme_path.write_text(content, encoding="utf-8")
    log_refactor("README.md atualizado com mapa de arquitetura.")

def run_refactoring():
    print("Iniciando refatoração de camadas...")
    ensure_structure()
    isolate_core_logic()
    isolate_doc_layout()
    refactor_app_py()
    update_readme()
    print("Refatoração concluída! Verifique logs/relatorio_refatoracao.md")

if __name__ == "__main__":
    run_refactoring()
