import unittest
from pathlib import Path
from core.validator import PromptValidator

class TestExplicacaoValidator(unittest.TestCase):
    def setUp(self):
        self.validator = PromptValidator()
        self.pre_correta = "Está correta. Se você marcou esta alternativa, provavelmente"
        self.pre_incorreta = "Está incorreta. Se você marcou esta alternativa, possivelmente"

    def test_sucesso_explicacao_consistente(self):
        dados = {
            'A': f"{self.pre_correta} entendeu o texto.",
            'B': f"{self.pre_incorreta} confundiu os termos.",
            'C': f"{self.pre_incorreta} errou a interpretação.",
            'D': f"{self.pre_incorreta} não leu o suporte.",
            'gabarito': 'A'
        }
        erros = self.validator.validar_explicacao_aluno(dados)
        self.assertEqual(len(erros), 0)

    def test_falha_nenhuma_correta(self):
        dados = {
            'A': f"{self.pre_incorreta} ...",
            'B': f"{self.pre_incorreta} ...",
            'C': f"{self.pre_incorreta} ...",
            'D': f"{self.pre_incorreta} ...",
            'gabarito': 'A'
        }
        erros = self.validator.validar_explicacao_aluno(dados)
        self.assertIn("Inconsistência na explicação ao aluno: nenhuma alternativa foi marcada como correta.", erros)

    def test_falha_multiplas_corretas(self):
        dados = {
            'A': f"{self.pre_correta} ...",
            'B': f"{self.pre_correta} ...",
            'C': f"{self.pre_incorreta} ...",
            'D': f"{self.pre_incorreta} ...",
            'gabarito': 'A'
        }
        erros = self.validator.validar_explicacao_aluno(dados)
        self.assertIn("Inconsistência na explicação ao aluno: há mais de uma alternativa marcada como correta.", erros)

    def test_falha_gabarito_divergente(self):
        dados = {
            'A': f"{self.pre_incorreta} ...",
            'B': f"{self.pre_correta} ...",
            'C': f"{self.pre_incorreta} ...",
            'D': f"{self.pre_incorreta} ...",
            'gabarito': 'C'
        }
        erros = self.validator.validar_explicacao_aluno(dados)
        self.assertIn("Inconsistência na explicação ao aluno: o gabarito informado não corresponde à alternativa marcada como correta.", erros)

if __name__ == "__main__":
    # Verificação de ambiente local
    root = Path(__file__).resolve().parent.parent
    if not (root / "core").exists() or not (root / "tests").exists():
        print("Erro: Estrutura de diretórios core/ ou tests/ não encontrada.")
    else:
        unittest.main()
