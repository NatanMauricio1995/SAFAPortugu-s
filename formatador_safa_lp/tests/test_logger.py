import unittest
import json
from pathlib import Path
from core.logger import SafaLogger

class TestSafaLogger(unittest.TestCase):
    def setUp(self):
        self.log_file = "logs/test_relatorio.json"
        self.logger = SafaLogger(self.log_file)
        self.full_path = Path("formatador_safa_lp") / self.log_file

    def tearDown(self):
        if self.full_path.exists():
            self.full_path.unlink()

    def test_sucesso_registro_padrao_safa(self):
        dados = {
            "acao": "padrao_safa",
            "total_itens_encontrados": 2,
            "itens_processados": [1, 2],
            "arquivos_gerados": ["output/item_01.docx", "output/item_02.docx"]
        }
        registro = self.logger.registrar_evento(dados)
        
        self.assertEqual(registro["acao"], "padrao_safa")
        self.assertEqual(len(registro["itens_processados"]), 2)
        
        with open(self.full_path, "r", encoding="utf-8") as f:
            historico = json.load(f)
            self.assertEqual(len(historico), 1)

    def test_falha_registro_erro_word(self):
        dados = {
            "acao": "conversao_png",
            "erros": [{"item": None, "mensagem": "Falha do Microsoft Word: Documento não encontrado."}]
        }
        self.logger.registrar_evento(dados)
        
        with open(self.full_path, "r", encoding="utf-8") as f:
            historico = json.load(f)
            self.assertEqual(historico[0]["acao"], "conversao_png")
            self.assertEqual(historico[0]["erros"][0]["mensagem"], "Falha do Microsoft Word: Documento não encontrado.")

if __name__ == "__main__":
    # Validação de ambiente
    base = Path("formatador_safa_lp")
    if not (base / "core").exists():
        print("Erro: Pasta core/ não encontrada.")
    else:
        unittest.main()
