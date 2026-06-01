import sys
import json
from pathlib import Path

def bootstrap():
    """
    Inicializa o sistema validando a integridade do ambiente e das camadas locais
    conforme as regras de isolamento e idempotência.
    """
    print("=" * 60)
    print("INICIALIZANDO FORMATADOR SAFA — LÍNGUA PORTUGUESA")
    print("=" * 60)
    
    # caminhos de controle baseados na raiz do projeto (onde o main.py reside)
    raiz = Path(__file__).parent.resolve()
    pastas_obrigatorias = ['core', 'data', 'templates', 'layout_interface', 'layout_documento', 'logs', 'tests']
    
    # 1. Validação física de diretórios
    for pasta in pastas_obrigatorias:
        caminho_pasta = raiz / pasta
        if not caminho_pasta.exists():
            print(f"[-] Erro Crítico: Diretório essencial '{pasta}/' ausente na raiz.")
            sys.exit(1)
            
    # 2. Leitura defensiva do arquivo de controle de versão
    caminho_versao = raiz / "version.json"
    if not caminho_versao.exists():
        print("[-] Erro Crítico: Arquivo 'version.json' não foi localizado.")
        sys.exit(1)
        
    try:
        with open(caminho_versao, 'r', encoding='utf-8') as f:
            v_info = json.load(f)
        print(f"[+] Sistema Carregado: {v_info.get('sistema')}")
        print(f"[+] Versão Core: {v_info.get('versao_sistema')} | Matrizes: {v_info.get('versao_matrizes')}")
    except json.JSONDecodeError:
        print("[-] Erro Crítico: O arquivo 'version.json' está corrompido ou malformado.")
        sys.exit(1)
        
    # 3. Inicialização e checagem do histórico central de logs
    caminho_log = raiz / "logs" / "relatorio_processamento.json"
    if not caminho_log.exists():
        log_inicial = [{
            "data_hora": "2026-06-01T12:00:00",
            "versao_sistema": v_info.get('versao_sistema'),
            "versao_matrizes": v_info.get('versao_matrizes'),
            "acao": "inicializacao_sistema",
            "status_inicializacao": "OK"
        }]
        with open(caminho_log, 'w', encoding='utf-8') as f:
            json.dump(log_inicial, f, indent=4, ensure_ascii=False)
        print("[+] Arquivo centralizado de telemetria inicializado em 'logs/'.")
        
    print("\n[✓] Todas as validações das camadas foram concluídas com sucesso.")
    print("[✓] O ambiente está pronto e isolado. Execute a interface ou os testes locais.")
    print("=" * 60)

if __name__ == "__main__":
    bootstrap()
