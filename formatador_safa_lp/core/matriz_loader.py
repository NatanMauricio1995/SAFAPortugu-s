import json
from pathlib import Path
from typing import Any, Dict, List


class MatrizLoader:
    """Carrega e consulta o arquivo data/matrizes_portugues.json.

    A resolução de caminho é feita de forma relativa à raiz do projeto, para
    funcionar tanto quando o programa é executado de dentro de formatador_safa_lp
    quanto a partir da pasta superior.
    """

    def __init__(self, file_path: str = "data/matrizes_portugues.json"):
        self.base_path = Path(__file__).resolve().parents[1]
        self.file_path = self.base_path / file_path
        self.data = self._load_json()

    def _load_json(self) -> Dict[str, Any]:
        if not self.file_path.exists():
            raise FileNotFoundError(f"Arquivo de matrizes não encontrado: {self.file_path}")
        try:
            with open(self.file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError as exc:
            raise ValueError("Não foi possível carregar o arquivo de matrizes. Verifique a estrutura do JSON.") from exc

    @staticmethod
    def _codigo(item: Dict[str, Any]) -> str:
        return str(item.get("código") or item.get("codigo") or "").strip()

    @staticmethod
    def _descricao(item: Dict[str, Any]) -> str:
        return str(item.get("descrição") or item.get("descricao") or "").strip()

    @staticmethod
    def normalizar_codigo(codigo: str) -> str:
        codigo = str(codigo or "").strip().upper().replace(" ", "")
        # Permite que o usuário digite D1 e encontre D01, D001 etc.
        if codigo.startswith("D"):
            numero = codigo[1:]
            if numero.isdigit():
                return f"D{int(numero)}"
        return codigo

    def get_matrizes_ativas(self) -> List[Dict[str, Any]]:
        ativas = [m for m in self.data.get("matrizes", []) if m.get("ativo", True)]
        return sorted(ativas, key=lambda x: x.get("ordem", 99))

    def get_etapas(self, matriz_id: str) -> List[Dict[str, Any]]:
        for matriz in self.data.get("matrizes", []):
            if matriz.get("id") == matriz_id:
                return matriz.get("etapas", [])
        return []

    def get_codigos(self, matriz_id: str, etapa_id: str) -> List[Dict[str, Any]]:
        etapas = self.get_etapas(matriz_id)
        for etapa in etapas:
            if etapa.get("id") == etapa_id:
                resultado = []
                for codigo in etapa.get("codigos", []):
                    cod = self._codigo(codigo)
                    desc = self._descricao(codigo)
                    display = f"{cod} - {desc}" if desc else cod
                    meta = dict(codigo)
                    meta.update({
                        "codigo_normalizado": self.normalizar_codigo(cod),
                        "codigo_exibicao": display,
                        "matriz_id": matriz_id,
                        "etapa_id": etapa_id,
                    })
                    resultado.append({"display": display, "meta": meta})
                return resultado
        return []

    def buscar_codigo(self, matriz_id: str, etapa_id: str, codigo_digitado: str) -> Dict[str, Any] | None:
        alvo = self.normalizar_codigo(codigo_digitado)
        for item in self.get_codigos(matriz_id, etapa_id):
            meta = item["meta"]
            cod = self._codigo(meta)
            if self.normalizar_codigo(cod) == alvo:
                return item
        return None
