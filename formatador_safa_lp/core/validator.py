from pathlib import Path

class PromptValidator:
    def validar(self, dados: dict) -> list[str]:
        """
        Valida o dicionário de dados do formulador.
        Retorna uma lista de strings contendo os erros encontrados.
        """
        erros = []
        
        # 1. Validações Gerais (Campos Obrigatórios)
        campos_gerais = {
            "modo_trabalho": "Informe o modo de trabalho.",
            "matriz_base": "Informe a matriz-base.",
            "etapa_matriz": "Informe a etapa da matriz.",
            "ano_alvo": "Informe o ano-alvo.",
            "codigo_alvo": "Informe o código-alvo.",
            "dificuldade": "Informe a dificuldade.",
            "tipo_suporte": "Informe o tipo de suporte."
        }
        
        for campo, mensagem in campos_gerais.items():
            if not dados.get(campo):
                erros.append(mensagem)
        
        # 2. Validações por Modo de Trabalho
        modo = dados.get("modo_trabalho")
        if modo == "Novo Item":
            if not dados.get("gabarito"):
                erros.append("Informe o gabarito desejado.")
            if not dados.get("item_base"):
                erros.append("Informe o item-base ou item de referência.")
        elif modo == "Revisão":
            if not dados.get("item_revisao"):
                erros.append("Informe o item para revisão.")
            # Gabarito em Revisão é opcional
        
        # 3. Validações Cruzadas
        # Texto-base sem fonte
        if dados.get("texto_base") and str(dados.get("texto_base")).strip():
            if not dados.get("fonte_autor") or not str(dados.get("fonte_autor")).strip():
                erros.append("Informe a fonte/autor do texto-base.")
        
        # Suporte visual sem descrição
        tipo_suporte = dados.get("tipo_suporte")
        if tipo_suporte and tipo_suporte != "Sem suporte":
            if not dados.get("descricao_suporte") or not str(dados.get("descricao_suporte")).strip():
                erros.append("O suporte selecionado exige descrição.")
                
        return erros
