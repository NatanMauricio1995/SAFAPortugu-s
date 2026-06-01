import unittest
from core.matriz_loader import MatrizLoader

class TestMatrizLoader(unittest.TestCase):
    def setUp(self):
        self.loader = MatrizLoader()

    def test_matrizes_ativas(self):
        ativas = self.loader.get_matrizes_ativas()
        self.assertTrue(len(ativas) >= 2)
        self.assertEqual(ativas[0]['id'], 'saeb_2001')

    def test_etapas_saeb_2001(self):
        etapas = self.loader.get_etapas("saeb_2001")
        ids = [e['id'] for e in etapas]
        self.assertIn("5_ano", ids)
        self.assertIn("9_ano", ids)

    def test_codigos_cascata(self):
        # Teste SAEB 2001
        res_2001 = self.loader.get_codigos("saeb_2001", "5_ano")
        self.assertEqual(res_2001[0]['meta']['código'], "D01")
        
        # Teste SAEB 2018
        res_2018 = self.loader.get_codigos("saeb_2018", "2_ano")
        self.assertEqual(res_2018[0]['meta']['código'], "EF12LP01")

if __name__ == "__main__":
    unittest.main()
