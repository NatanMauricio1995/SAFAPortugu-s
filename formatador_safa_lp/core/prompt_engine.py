import os
from pathlib import Path
from datetime import datetime

class PromptEngine:
    def __init__(self, templates_dir="templates"):
        self.base_path = Path("formatador_safa_lp")
        self.templates_dir = self.base_path / templates_dir

    def gerar_prompt(self, dados: dict) -> str:
        modo = dados.get('modo_trabalho', 'Novo Item')
        template_name = "prompt_portugues_novo_item.txt" if modo == "Novo Item" else "prompt_portugues_revisao_item.txt"
        template_path = self.templates_dir / template_name

        if not template_path.exists():
            raise FileNotFoundError(f"Template não encontrado: {template_path}")

        template_content = template_path.read_text(encoding="utf-8")
        
        # Lógica de Adaptação Controlada
        if str(dados.get('ano_alvo')) != str(dados.get('etapa_matriz')):
            dados['observacao_adaptacao'] = "Observação de adaptação: o ano-alvo é diferente da etapa da matriz. A resposta deve aplicar adaptação controlada de linguagem, suporte e complexidade, sem alterar o foco do código-alvo."
        else:
            dados['observacao_adaptacao'] = "Observação de adaptação: não se aplica."

        # Tratamento de Vazios Opcionais
        mapeamento_vazios = {
            'tema_contexto': 'não informado',
            'texto_base': 'não informado',
            'restricoes': 'não informadas',
            'fonte_autor': 'não se aplica'
        }
        
        for chave, valor_padrao in mapeamento_vazios.items():
            if not dados.get(chave):
                dados[chave] = valor_padrao

        # Placeholders obrigatórios fixos/gerados
        dados.setdefault('versao_template', '1.0.0')
        dados.setdefault('data_geracao', datetime.now().strftime("%d/%m/%Y %H:%M"))
        dados.setdefault('componente', 'Língua Portuguesa')
        dados.setdefault('item_base_ou_revisao', dados.get('texto_base', ''))

        # Renderização
        prompt = template_content
        for chave, valor in dados.items():
            placeholder = f"{{{{{chave}}}}}"
            prompt = prompt.replace(placeholder, str(valor))

        return prompt
