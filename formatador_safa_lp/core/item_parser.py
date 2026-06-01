import re
from pathlib import Path

class ItemParser:
    def __init__(self):
        self.item_marker_regex = re.compile(r'^#\s*Item\s+(\d+)', re.MULTILINE)
        self.field_map = {
            'texto_base': r'Texto-base:\s*(.*?)(?=\n\w+:|\n[A-D]\)|$)',
            'fonte_autor': r'Fonte:\s*(.*?)(?=\n\w+:|\n[A-D]\)|$)',
            'suporte': r'Suporte:\s*(.*?)(?=\n\w+:|\n[A-D]\)|$)',
            'enunciado': r'Enunciado:\s*(.*?)(?=\n\w+:|\n[A-D]\)|$)',
            'comando': r'Comando:\s*(.*?)(?=\n\w+:|\n[A-D]\)|$)',
            'alternativa_a': r'A\)\s*(.*?)(?=\n[B-D]\)|\nGabarito:|$)',
            'alternativa_b': r'B\)\s*(.*?)(?=\n[C-D]\)|\nGabarito:|$)',
            'alternativa_c': r'C\)\s*(.*?)(?=\nD\)\|\nGabarito:|$)',
            'alternativa_d': r'D\)\s*(.*?)(?=\nGabarito:|$)',
            'gabarito': r'Gabarito:\s*(.*?)(?=\n\w+:|$)',
            'explicacao_aluno': r'Explicação ao aluno:\s*(.*?)(?=\n\w+:|$)'
        }

    def separar_e_parsear(self, texto_lote: str) -> list[dict]:
        matches = list(self.item_marker_regex.finditer(texto_lote))
        
        if not matches:
            raise ValueError("Nenhum item foi encontrado. Use o marcador # Item 1, # Item 2, etc.")
        
        ids = [int(m.group(1)) for m in matches]
        
        # Validação de IDs
        if len(ids) != len(set(ids)):
            raise ValueError("Há itens com numeração repetida.")
        
        if sorted(ids) != list(range(min(ids), max(ids) + 1)):
            print("A numeração dos itens não está sequencial.")
            
        itens_processados = []
        for i, match in enumerate(matches):
            start = match.end()
            end = matches[i+1].start() if i + 1 < len(matches) else len(texto_lote)
            conteudo = texto_lote[start:end].strip()
            
            dados_item = {
                'numero': ids[i],
                'conteudo_bruto': conteudo
            }
            
            for campo, regex in self.field_map.items():
                match_campo = re.search(regex, conteudo, re.DOTALL | re.IGNORECASE)
                dados_item[campo] = match_campo.group(1).strip() if match_campo else ""
                
            itens_processados.append(dados_item)
            
        return itens_processados
