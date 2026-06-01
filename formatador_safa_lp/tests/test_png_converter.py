import unittest
from pathlib import Path
from core.png_converter import PngConverter

class TestPngConverter(unittest.TestCase):
    def setUp(self):
        self.converter = PngConverter()
        self.base = Path("formatador_safa_lp")
        # Nota: Este teste exige Word instalado e um arquivo .docx real para passar integralmente.
        # Aqui simulamos a estrutura de caminhos.

    def test_verificacao_pastas(self):
        self.assertTrue((self.base / "core").exists())
        self.assertTrue((self.base / "tests").exists())

    def test_mock_conversao_falha_arquivo_inexistente(self):
        resultado = self.converter.converter_documento("arquivo_fantasma.docx")
        self.assertEqual(resultado["status"], "erro")
        self.assertIn("não encontrado", resultado["erro"])

    # Testes funcionais reais requerem ambiente Windows com Word
    # def test_conversao_real(self):
    #     docx_teste = self.base / "output" / "item_01_padronizado.docx"
    #     if docx_teste.exists():
    #         res = self.converter.converter_documento(str(docx_teste), cortar_conteudo=True)
    #         self.assertEqual(res["status"], "sucesso")
    #         self.assertTrue(len(res["imagens"]) > 0)

if __name__ == "__main__":
    unittest.main()
