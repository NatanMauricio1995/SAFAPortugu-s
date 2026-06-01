from pathlib import Path
from datetime import datetime


class PromptEngine:
    def __init__(self, templates_dir: str = "templates"):
        self.base_path = Path(__file__).resolve().parents[1]
        self.templates_dir = self.base_path / templates_dir

    def gerar_prompt(self, dados: dict) -> str:
        modo = str(dados.get("modo_trabalho", "Novo Item"))
        template_name = "prompt_portugues_revisao_item.txt" if "revis" in modo.lower() else "prompt_portugues_novo_item.txt"
        template_path = self.templates_dir / template_name

        if not template_path.exists():
            raise FileNotFoundError(f"Template não encontrado: {template_path}")

        dados = dict(dados)
        template_content = template_path.read_text(encoding="utf-8")

        if str(dados.get("ano_alvo", "")).strip() != str(dados.get("etapa_matriz", "")).strip():
            dados["observacao_adaptacao"] = (
                "Observação de adaptação: o ano-alvo é diferente da etapa da matriz. "
                "A resposta deve aplicar adaptação controlada de linguagem, suporte e complexidade, "
                "sem alterar o foco do código-alvo."
            )
        else:
            dados["observacao_adaptacao"] = "Observação de adaptação: não se aplica."

        mapeamento_vazios = {
            "tema_contexto": "não informado",
            "tipo_suporte": "não informado",
            "descricao_suporte": "não se aplica",
            "texto_base": "não informado",
            "restricoes": "não informadas",
            "fonte_autor": "não se aplica",
            "descricao_codigo": "não informada",
            "gabarito": "não informado",
            "item_base_ou_revisao": "não informado",
        }

        for chave, valor_padrao in mapeamento_vazios.items():
            if not dados.get(chave):
                dados[chave] = valor_padrao

        dados.setdefault("versao_template", "1.0.0")
        dados.setdefault("data_geracao", datetime.now().strftime("%d/%m/%Y %H:%M"))
        dados.setdefault("componente", "Língua Portuguesa")

        prompt = template_content
        for chave, valor in dados.items():
            prompt = prompt.replace(f"{{{{{chave}}}}}", str(valor))

        return prompt
