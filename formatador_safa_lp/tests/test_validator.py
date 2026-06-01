import unittest
from pathlib import Path
from core.validator import PromptValidator

class TestPromptValidator(unittest.TestCase):
    def setUp(self):
        self.validator = PromptValidator()

    def test_sucesso_novo_item(self):
        dados = {
            "modo_trabalho": "Novo Item",
            "matriz_base": "SAEB 2018",
            "etapa_matriz": "5º Ano",
            "ano_alvo": "5º Ano",
            "codigo_alvo": "EF15LP03",
            "dificuldade": "Média",
            "tipo_suporte": "Sem suporte",
            "gabarito": "A",
            "item_base": "Enunciado de teste"
        }
        erros = self.validator.validar(dados)
        self.assertEqual(len(erros), 0)

    def test_sucesso_revisao(self):
        dados = {
            "modo_trabalho": "Revisão",
            "matriz_base": "SAEB 2018",
            "etapa_matriz": "9º Ano",
            "ano_alvo": "9º Ano",
            "codigo_alvo": "EF69LP03",
            "dificuldade": "Difícil",
            "tipo_suporte": "Sem suporte",
            "item_revisao": "Item original para revisar"
        }
        erros = self.validator.validar(dados)
        self.assertEqual(len(erros), 0)

    def test_falha_campos_gerais(self):
        dados = {}
        erros = self.validator.validar(dados)
        self.assertIn("Informe o modo de trabalho.", erros)
        self.assertIn("Informe a matriz-base.", erros)
        self.assertIn("Informe o tipo de suporte.", erros)

    def test_falha_novo_item_incompleto(self):
        dados = {"modo_trabalho": "Novo Item"}
        erros = self.validator.validar(dados)
        self.assertIn("Informe o gabarito desejado.", erros)
        self.assertIn("Informe o item-base ou item de referência.", erros)

    def test_falha_revisao_incompleta(self):
        dados = {"modo_trabalho": "Revisão"}
        erros = self.validator.validar(dados)
        self.assertIn("Informe o item para revisão.", erros)

    def test_falha_cruzada_fonte_texto(self):
        dados = {
            "modo_trabalho": "Novo Item", "matriz_base": "X", "etapa_matriz": "X",
            "ano_alvo": "X", "codigo_alvo": "X", "dificuldade": "X",
            "tipo_suporte": "Sem suporte", "gabarito": "A", "item_base": "X",
            "texto_base": "Texto longo sem fonte"
        }
        erros = self.validator.validar(dados)
        self.assertIn("Informe a fonte/autor do texto-base.", erros)

    def test_falha_cruzada_suporte_descricao(self):
        dados = {
            "modo_trabalho": "Novo Item", "matriz_base": "X", "etapa_matriz": "X",
            "ano_alvo": "X", "codigo_alvo": "X", "dificuldade": "X",
            "tipo_suporte": "Imagem", "gabarito": "A", "item_base": "X"
        }
        erros = self.validator.validar(dados)
        self.assertIn("O suporte selecionado exige descrição.", erros)

if __name__ == "__main__":
    # Garantir que o ambiente de execução é válido antes de prosseguir
    project_root = Path(__file__).resolve().parent.parent
    if not (project_root / "core").exists() or not (project_root / "tests").exists():
        print("Erro: Estrutura de pastas core/ ou tests/ não encontrada.")
    else:
        unittest.main()
