import unittest
import sys
from datetime import datetime
from pathlib import Path
from test_suite_safa import TestSuiteSafa

def gerar_relatorio():
    loader = unittest.TestLoader()
    suite = loader.loadTestsFromTestCase(TestSuiteSafa)
    
    result = unittest.TestResult()
    suite.run(result)
    
    base_path = Path("formatador_safa_lp")
    relatorio_path = base_path / "tests" / "RELATORIO_TESTES.md"
    
    status_geral = "Sucesso" if result.wasSuccessful() else "Falhas Encontradas"
    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    
    md_content = f"""# Relatório de Homologação — Formatador SAFA LP

- **Data/Hora Execution:** {timestamp}
- **Status Geral:** {status_geral}

## Tabela de Testes
| Caso de Teste | Status | Evidência/Erro Encontrado |
| :--- | :--- | :--- |
"""
    
    # Processar Sucessos (Simulado, unittest.TestResult não lista sucessos por padrão sem Custom Class)
    # Aqui listamos todos os testes e marcamos o status
    all_tests = [t._testMethodName for t in suite]
    failures = {f[0]._testMethodName: f[1] for f in result.failures}
    errors = {e[0]._testMethodName: e[1] for e in result.errors}

    for test_name in all_tests:
        status = "✅ PASS"
        evidencia = "-"
        if test_name in failures:
            status = "❌ FAIL"
            evidencia = failures[test_name].split('\n')[-2]
        elif test_name in errors:
            status = "⚠️ ERROR"
            evidencia = errors[test_name].split('\n')[-2]
            
        md_content += f"| {test_name} | {status} | {evidencia} |\n"

    md_content += "\n## Ajustes e Correções\n"
    if not result.wasSuccessful():
        md_content += "- [ ] Revisar falhas reportadas na tabela acima.\n"
    else:
        md_content += "- [x] Todos os módulos core validados e isolados do layout.\n"

    relatorio_path.write_text(md_content, encoding="utf-8")
    print(f"Relatório gerado em: {relatorio_path}")

if __name__ == "__main__":
    gerar_relatorio()
