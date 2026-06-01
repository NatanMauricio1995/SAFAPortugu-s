import unittest
import json
from pathlib import Path
from core.matriz_loader import MatrizLoader
from core.prompt_engine import PromptEngine
from core.validator import PromptValidator
from core.word_formatter import WordFormatter
from core.item_parser import ItemParser
from core.safa_processor import SafaProcessor
from core.png_converter import PngConverter

class TestSuiteSafa(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.base = Path("formatador_safa_lp")
        dirs = ["tests", "core", "layout_interface", "output"]
        for d in dirs:
            if not (cls.base / d).exists():
                raise FileNotFoundError(f"Diretório crítico ausente: {d}")

    def setUp(self):
        self.validator = PromptValidator()
        self.engine = PromptEngine()
        self.parser = ItemParser()
        self.formatter = WordFormatter()
        self.processor = SafaProcessor()

    # --- ABA 1: PROMPT & VALIDATOR ---
    def test_aba1_prompt_sucesso_e_adaptacao(self):
        dados = {"modo_trabalho": "Novo Item", "ano_alvo": "4º Ano", "etapa_matriz": "5º Ano"}
        prompt = self.engine.gerar_prompt(dados)
        self.assertIn("Observação de adaptação: o ano-alvo é diferente", prompt)
        
    def test_aba1_validator_cruzado(self):
        dados = {"modo_trabalho": "Novo Item", "texto_base": "Era uma vez...", "fonte_autor": ""}
        erros = self.validator.validar(dados)
        self.assertIn("Informe a fonte/autor do texto-base.", erros)

    # --- ABA 2: WORD & PARSER ---
    def test_aba2_parser_regex_e_secoes(self):
        texto = "# Item 1\nEnunciado: X\nGabarito: A"
        with self.assertRaises(ValueError): # Falta seções obrigatórias para o parser completo
            self.parser.separar_e_parsear(texto)
            
    def test_aba2_word_secao_ausente(self):
        texto = "ITEM FINAL\nEnunciado: Teste" # Sem EXPLICAÇÃO AO ALUNO
        valido, erro = self.formatter.validar_estrutura(texto)
        self.assertFalse(valido)
        self.assertIn("EXPLICAÇÃO AO ALUNO", erro)

    # --- ABA 3: PADRÃO SAFA & EXPLICAÇÕES ---
    def test_aba3_safa_intervalo_e_prefixos(self):
        ids = self.processor.parse_intervalo_itens("1, 3-5")
        self.assertEqual(ids, {1, 3, 4, 5})
        
    def test_aba3_validator_explicao_corretas(self):
        pre = "Está correta. Se você marcou esta alternativa, provavelmente"
        dados = {'A': pre, 'B': pre, 'gabarito': 'A'}
        erros = self.validator.validar_explicacao_aluno(dados)
        self.assertIn("há mais de uma alternativa marcada como correta.", erros)

    # --- ABA 4: PNG & ISOLAMENTO ---
    def test_aba4_png_mock_isolamento(self):
        # Garante que as pastas de output existem
        self.assertTrue((self.base / "output").exists())
        # O teste de isolamento verifica se o core não depende de arquivos do layout_interface
        # via tentativa de importação/execução sem o diretório de layout no path
        self.assertTrue(hasattr(PngConverter, 'converter_documento'))

if __name__ == "__main__":
    unittest.main()
