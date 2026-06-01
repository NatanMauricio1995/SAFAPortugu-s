import json
import threading
from pathlib import Path
from datetime import datetime

class SafaLogger:
    def __init__(self, log_path="logs/relatorio_processamento.json"):
        self.base_path = Path("formatador_safa_lp")
        self.file_path = self.base_path / log_path
        self._lock = threading.Lock()
        self._inicializar_log()

    def _inicializar_log(self):
        log_dir = self.file_path.parent
        if not log_dir.exists():
            log_dir.mkdir(parents=True, exist_ok=True)
        
        if not self.file_path.exists():
            with open(self.file_path, "w", encoding="utf-8") as f:
                json.dump([], f)

    def registrar_evento(self, dados_evento: dict):
        """
        Registra um evento de processamento de forma atômica no histórico JSON.
        """
        with self._lock:
            # Carrega histórico atual
            try:
                with open(self.file_path, "r", encoding="utf-8") as f:
                    historico = json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                historico = []

            # Monta o novo registro baseado no esquema obrigatório
            novo_registro = {
                "data_hora": datetime.now().isoformat(),
                "versao_sistema": dados_evento.get("versao_sistema", "1.0.0"),
                "versao_matrizes": dados_evento.get("versao_matrizes", "1.0.0"),
                "versao_template_novo_item": dados_evento.get("versao_template_novo_item", "1.0.0"),
                "versao_template_revisao": dados_evento.get("versao_template_revisao", "1.0.0"),
                "acao": dados_evento.get("acao", "erro_sistema"),
                "arquivo_origem": dados_evento.get("arquivo_origem", ""),
                "total_itens_encontrados": int(dados_evento.get("total_itens_encontrados", 0)),
                "itens_processados": dados_evento.get("itens_processados", []),
                "itens_ignorados": dados_evento.get("itens_ignorados", []),
                "itens_com_erro": dados_evento.get("itens_com_erro", []),
                "erros": dados_evento.get("erros", []), # [{'item': int/None, 'mensagem': str}]
                "arquivos_gerados": dados_evento.get("arquivos_gerados", []),
                "imagens_geradas": dados_evento.get("imagens_geradas", [])
            }

            historico.append(novo_registro)

            # Salva atômicamente
            with open(self.file_path, "w", encoding="utf-8") as f:
                json.dump(historico, f, indent=4, ensure_ascii=False)
            
            return novo_registro
