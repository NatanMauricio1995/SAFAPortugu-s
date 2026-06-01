import unittest
from pathlib import Path
from core.safa_processor import SafaProcessor

class TestSafaProcessor(unittest.TestCase):
    def setUp(self):
        self.processor = SafaProcessor()

    def test_parse_intervalo_complexo(self):
        entrada = "1, 3, 8-10; 12"
        esperado = {1, 3, 8, 9, 10, 12}
        self.assertEqual(self.processor.parse_intervalo_itens(entrada), esperado)

    def test_geracao_lote_6_paginas(self):
        base = Path("formatador_safa_lp")
        output_test = base / "output" / "test_batch"
        
        itens = [{
            'numero': 1,
            'enunciado': 'Enunciado teste',
            'comando': 'Comando teste',
            'alternativa_a': 'A) Alt A',
            'alternativa_b': 'B) Alt B',
            'alternativa_c': 'C) Alt C',
            'alternativa_d': 'D) Alt D',
            'gabarito': 'A',
            'explicacao_aluno': 'Explicação teste'
        }]
        
        configs = {
            'pasta_saida': output_test,
            'incluir_ensinart': True
        }
        
        relatorio = self.processor.processar_lote(itens, {1}, configs)
        
        self.assertEqual(relatorio['itens'][0]['status'], 'sucesso')
        arquivo_gerado = output_test / "item_01_padronizado.docx"
        self.assertTrue(arquivo_gerado.exists())
        
        # Validação do relatório
        relatorio_file = base / "logs" / "relatorio_processamento.json"
        self.assertTrue(relatorio_file.exists())

    def test_validacao_item_invalido(self):
        itens = [{'numero': 5, 'enunciado': 'Incompleto'}]
        relatorio = self.processor.processar_lote(itens, set(), {'pasta_saida': 'output'})
        self.assertEqual(relatorio['itens'][0]['status'], 'erro')
        self.assertTrue(len(relatorio['itens'][0]['erros']) > 0)

if __name__ == "__main__":
    unittest.main()
