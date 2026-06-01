import unittest
from pathlib import Path
from core.item_parser import ItemParser

class TestItemParser(unittest.TestCase):
    def setUp(self):
        self.parser = ItemParser()

    def test_parse_sucesso_sequencial(self):
        texto = """
# Item 1
Enunciado: Questão 1
A) Opção A
B) Opção B
Gabarito: A
# Item 2
Enunciado: Questão 2
Gabarito: B
"""
        itens = self.parser.separar_e_parsear(texto)
        self.assertEqual(len(itens), 2)
        self.assertEqual(itens[0]['numero'], 1)
        self.assertEqual(itens[1]['numero'], 2)

    def test_falha_sem_marcadores(self):
        with self.assertRaisesRegex(ValueError, "Nenhum item foi encontrado"):
            self.parser.separar_e_parsear("Texto sem marcadores")

    def test_falha_numeracao_repetida(self):
        texto = "# Item 1\n# Item 1"
        with self.assertRaisesRegex(ValueError, "numeração repetida"):
            self.parser.separar_e_parsear(texto)

    def test_aviso_numeracao_pulada(self):
        # Deve funcionar mas imprimir aviso (simulado aqui apenas pela conclusão do parse)
        texto = "# Item 1\n# Item 3"
        itens = self.parser.separar_e_parsear(texto)
        self.assertEqual(len(itens), 2)
        self.assertEqual(itens[1]['numero'], 3)

if __name__ == "__main__":
    # Validação de ambiente
    root = Path(__file__).resolve().parent.parent
    if not (root / "core").exists() or not (root / "tests").exists():
        print("Erro: Estrutura core/ ou tests/ ausente.")
    else:
        unittest.main()
