import unittest
from pathlib import Path
from core.word_formatter import WordFormatter

class TestWordFormatter(unittest.TestCase):
    def setUp(self):
        self.formatter = WordFormatter()
        self.texto_completo = """
        MINI RELATÓRIO FINAL
        TABELA DE CHECAGEM TÉCNICA
        ITEM FINAL
        Texto-base: Texto de exemplo.
        Fonte: Autor Desconhecido.
        Enunciado: Questão teste.
        Comando: Escolha a correta.
        Alternativas
        A) Opção 1
        B) Opção 2
        C) Opção 3
        D) Opção 4
        Gabarito: A
        JUSTIFICATIVA DO GABARITO: Porque sim.
        JUSTIFICATIVA TÉCNICA DOS DISTRATORES: Porque não.
        EXPLICAÇÃO AO ALUNO: Veja bem...
        """

    def test_geracao_sucesso_com_fonte(self):
        resultado = self.formatter.gerar_docx(self.texto_completo, "teste_sucesso.docx")
        self.assertTrue("teste_sucesso.docx" in resultado)
        self.assertTrue(Path(resultado).exists())

    def test_geracao_sucesso_sem_texto_base(self):
        texto = self.texto_completo.replace("Texto-base: Texto de exemplo.", "").replace("Fonte: Autor Desconhecido.", "")
        resultado = self.formatter.gerar_docx(texto, "teste_sem_texto_base.docx")
        self.assertTrue("teste_sem_texto_base.docx" in resultado)

    def test_falha_secao_ausente(self):
        texto_incompleto = "ITEM FINAL\nEnunciado: ..."
        valido, erro = self.formatter.validar_estrutura(texto_incompleto)
        self.assertFalse(valido)
        self.assertIn("A seção 'MINI RELATÓRIO FINAL' não foi encontrada", erro)

    def test_falha_fonte_ausente(self):
        texto = self.texto_completo.replace("Fonte: Autor Desconhecido.", "")
        valido, erro = self.formatter.validar_estrutura(texto)
        self.assertFalse(valido)
        self.assertIn("seção 'Fonte' não foi encontrada", erro)

if __name__ == "__main__":
    unittest.main()
