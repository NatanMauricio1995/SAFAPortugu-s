import os
import shutil
import sys
import traceback
from pathlib import Path

# CONFIGURAÇÕES
SOURCE_DIR = Path("formatador_matematica")
DEST_DIR = Path("formatador_safa_lp")
BACKUP_DIR = DEST_DIR / "_backup_pre_migracao"

CORE_MODULES = [
    "matriz_loader.py", "prompt_engine.py", "validator.py", 
    "item_parser.py", "word_formatter.py", "safa_processor.py", 
    "png_converter.py", "logger.py"
]

FOLDERS = [
    "core", "data", "templates", "layout_interface", 
    "layout_documento", "output", "logs", "tests"
]

def setup_logs():
    log_dir = DEST_DIR / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # Diagnóstico Inicial
    diag_path = log_dir / "diagnostico_inicial.md"
    with open(diag_path, "w", encoding="utf-8") as f:
        f.write("# Diagnóstico Inicial de Migração\n\n")
        if DEST_DIR.exists():
            f.write(f"Destino '{DEST_DIR}' já existe. Conteúdo atual:\n")
            for item in DEST_DIR.rglob("*"):
                f.write(f"- {item.relative_to(DEST_DIR)}\n")
        else:
            f.write(f"Destino '{DEST_DIR}' não existe. Será criado do zero.\n")

def log_migration(message):
    log_path = DEST_DIR / "logs" / "registro_migracao_inicial.md"
    mode = "a" if log_path.exists() else "w"
    with open(log_path, mode, encoding="utf-8") as f:
        if mode == "w":
            f.write("# Registro de Migração Inicial\n\n")
        f.write(f"- {message}\n")

def migrate():
    # 1. Validação de Segurança
    if not SOURCE_DIR.exists():
        print(f"ERRO: Pasta de origem '{SOURCE_DIR}' não encontrada.")
        sys.exit(1)

    if not DEST_DIR.exists():
        DEST_DIR.mkdir(parents=True)
        log_migration("Pasta de destino criada.")
    else:
        # Idempotência: Backup de arquivos inesperados
        expected_items = set(FOLDERS) | {"app.py", "main.py", "requirements.txt", "README_IMPLEMENTACAO.md", "_backup_pre_migracao", "logs"}
        for item in DEST_DIR.iterdir():
            if item.name not in expected_items and item.name != "_backup_pre_migracao":
                BACKUP_DIR.mkdir(exist_ok=True)
                shutil.move(str(item), str(BACKUP_DIR / item.name))
                log_migration(f"Movido para backup: {item.name}")

    setup_logs()

    # 2. Estrutura de Pastas
    for folder in FOLDERS:
        (DEST_DIR / folder).mkdir(exist_ok=True)
        log_migration(f"Diretório verificado/criado: {folder}")

    # 3. READMEs e Stubs
    readme_map = {
        "data/README_DATA.md": "Repositório de matrizes e dados JSON de Língua Portuguesa.",
        "templates/README_TEMPLATES.md": "Templates de prompts para LLM (Novo Item / Revisão).",
        "layout_documento/README_LAYOUT_DOCUMENTO.md": "Configurações de estilos e layout .docx.",
        "README_IMPLEMENTACAO.md": "# Projeto Formatador SAFA — Língua Portuguesa\nInicializado via migrar_projeto.py"
    }
    
    for path, content in readme_map.items():
        file_path = DEST_DIR / path
        if not file_path.exists():
            file_path.write_text(content, encoding="utf-8")
            log_migration(f"Gerado: {path}")

    for mod in CORE_MODULES:
        mod_path = DEST_DIR / "core" / mod
        if not mod_path.exists():
            content = f'"""\nMódulo: {mod.replace(".py", "").replace("_", " ").title()}\nResponsável pela lógica de {mod.split(".")[0]}.\n"""\n\ndef init():\n    pass\n'
            mod_path.write_text(content, encoding="utf-8")
            log_migration(f"Stub criado: core/{mod}")

    # 4. Migração de Arquivos Base (app.py / main.py)
    source_app = SOURCE_DIR / "app.py"
    dest_app = DEST_DIR / "app.py"
    
    if source_app.exists() and not dest_app.exists():
        content = source_app.read_text(encoding="utf-8")
        # Substituição de strings
        content = content.replace("Matemática", "Língua Portuguesa")
        content = content.replace("matematica", "safa_lp")
        dest_app.write_text(content, encoding="utf-8")
        log_migration("app.py migrado com substituições de 'Matemática' para 'Língua Portuguesa'.")

    # 5. Interface (Herança de CSS/JS limpos)
    # Procura arquivos na raiz ou assets do projeto original
    assets_map = {
        "layout.html": "layout_interface/index.html",
        "assets/style.css": "layout_interface/style.css",
        "assets/app.js": "layout_interface/scripts.js"
    }

    for src_rel, dest_rel in assets_map.items():
        src_path = SOURCE_DIR / src_rel
        target_path = DEST_DIR / dest_rel
        if src_path.exists() and not target_path.exists():
            shutil.copy(src_path, target_path)
            log_migration(f"Arquivo visual herdado: {dest_rel}")

    # 6. Requirements
    req_path = DEST_DIR / "requirements.txt"
    if not req_path.exists():
        req_path.write_text("flask\npython-docx\npillow\nrequests\n", encoding="utf-8")
        log_migration("requirements.txt inicial gerado.")

    # 7. Validação (Dry-run)
    validate_migration()

def validate_migration():
    log_migration("Iniciando Validação (Dry-run)...")
    try:
        # Adiciona o diretório ao path para simular importação
        sys.path.append(str(DEST_DIR))
        
        # Tenta importar os módulos principais
        import app
        log_migration("SUCESSO: app.py importado corretamente.")
        
        from core import safa_processor
        log_migration("SUCESSO: core.safa_processor importado corretamente.")
        
    except Exception:
        error_msg = traceback.format_exc()
        log_migration(f"ERRO CRÍTICO NA VALIDAÇÃO:\n{error_msg}")
        print("Validação falhou. Verifique os logs em formatador_safa_lp/logs/.")
    else:
        log_migration("Validação concluída com sucesso.")
        print("Migração concluída com sucesso!")

if __name__ == "__main__":
    migrate()
