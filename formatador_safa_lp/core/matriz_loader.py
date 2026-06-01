import json
from pathlib import Path

class MatrizLoader:
    def __init__(self, file_path="data/matrizes_portugues.json"):
        self.base_path = Path("formatador_safa_lp")
        self.file_path = self.base_path / file_path
        self.data = self._load_json()

    def _load_json(self):
        if not self.file_path.exists():
            raise FileNotFoundError(f"Arquivo de matrizes não encontrado: {self.file_path}")
        try:
            with open(self.file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            raise ValueError("Não foi possível carregar o arquivo de matrizes. Verifique a estrutura do JSON.")

    def get_matrizes_ativas(self):
        ativas = [m for m in self.data.get("matrizes", []) if m.get("ativo")]
        return sorted(ativas, key=lambda x: x.get("ordem", 99))

    def get_etapas(self, matriz_id):
        for m in self.data.get("matrizes", []):
            if m["id"] == matriz_id:
                return m.get("etapas", [])
        return []

    def get_codigos(self, matriz_id, etapa_id):
        etapas = self.get_etapas(matriz_id)
        for e in etapas:
            if e["id"] == etapa_id:
                codigos = e.get("codigos", [])
                if not codigos:
                    raise Exception("Nenhum código ativo encontrado para a matriz e etapa selecionadas.")
                
                result = []
                for c in codigos:
                    formatted = f"{c['código']} - {c['descrição']}"
                    full_meta = c.copy()
                    full_meta.update({"matriz": matriz_id, "etapa": etapa_id})
                    result.append({"display": formatted, "meta": full_meta})
                return result
        raise Exception("Nenhum código ativo encontrado para a matriz e etapa selecionadas.")
