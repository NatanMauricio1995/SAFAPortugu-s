import os
from pathlib import Path
from core.prompt_engine import PromptEngine

def test_validation():
    # Setup de pastas
    base = Path("formatador_safa_lp")
    test_dir = base / "tests" / "exemplos_prompt"
    test_dir.mkdir(parents=True, exist_ok=True)

    engine = PromptEngine()

    dados_teste = {
        "modo_trabalho": "Novo Item",
        "matriz_base": "SAEB 2018",
        "etapa_matriz": "5º Ano",
        "ano_alvo": "4º Ano", # Diferente para testar adaptação
        "codigo_alvo": "EF15LP03",
        "descricao_codigo": "Localizar informações explícitas em textos.",
        "dificuldade": "Média",
        "tema_contexto": "Contos de fadas",
        "tipo_suporte": "Texto Narrativo",
        "descricao_suporte": "Um pequeno trecho da Cinderela",
        "texto_base": "", # Testar tratamento de vazio
        "fonte_autor": "",
        "restricoes": "Evitar termos arcaicos",
        "gabarito": "A"
    }

    try:
        prompt_gerado = engine.gerar_prompt(dados_teste)
        
        output_file = test_dir / "exemplo_gerado.txt"
        output_file.write_text(prompt_gerado, encoding="utf-8")
        
        print(f"Sucesso! Exemplo gerado em: {output_file}")
        
        # Validação básica de placeholders obrigatórios
        assert "Observação de adaptação:" in prompt_gerado
        assert "não informado" in prompt_gerado
        assert "EF15LP03" in prompt_gerado
        
    except Exception as e:
        print(f"Erro na validação: {e}")

if __name__ == "__main__":
    test_validation()
