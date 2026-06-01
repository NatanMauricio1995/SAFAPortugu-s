# Formatador SAFA — Língua Portuguesa (v1.0.0)

## 1. Objetivo do Programa
O **Formatador SAFA LP** é uma aplicação desktop/local voltada para engenharia pedagógica e automação de documentos educacionais de Língua Portuguesa. O sistema automatiza três macroprocessos:
1. **Montagem de Prompts Estruturados**: Formulação direcionada para criação ou revisão de itens com base em matrizes de proficiência (SAEB 2001/2018).
2. **Sanitização e Parseamento**: Segmentação de lotes de itens utilizando o marcador rígido `# Item X`.
3. **Diagramação Documental e Visual**: Geração de arquivos Word (`.docx`) com paginação fixa padrão SAFA (6 páginas por item) e extração de páginas em alta resolução (`.png`) via automação nativa do Microsoft Word.

---

## 2. Requisitos de Instalação (Modelo de Distribuição)
O sistema foi homologado estritamente para o ecossistema Windows, espelhando as diretrizes operacionais do módulo de Matemática.

### Pré-requisitos do Ambiente:
* **Sistema Operacional**: Windows 10 ou superior (64-bit).
* **Dependência de Software**: Microsoft Word instalado localmente e devidamente licenciado (necessário suporte a vínculos COM/OLE `win32com`).
* **Dependência do Sistema**: Poppler instalado e mapeado no PATH do Windows (exigido pela biblioteca `pdf2image` para conversão de PDF para PNG).
* **Ambiente de Execução**: Python 3.10+ instalado.

### Procedimento de Instalação Local:
1. Extraia o pacote compactado `formatador_safa_lp_v1.0.0.zip` na raiz do diretório desejado (ex: `C:\formatador_safa_lp`).
2. Abra o terminal (PowerShell ou Prompt de Comando) na pasta do projeto.
3. Instale as dependências via gerenciador de pacotes pip:
   ```bash
   pip install -r requirements.txt
   ```

4. Execute o ponto de entrada principal do sistema (localizado no diretório raiz):
```bash
python main.py
```

---

## 3. Arquitetura do Projeto e Estrutura de Pastas

O projeto adota o princípio de **Isolamento de Responsabilidades (Layered Architecture)**. O núcleo algorítmico é estritamente desacoplado das parametrizações visuais da interface gráfica e do comportamento geométrico dos documentos impressos.

```text
formatador_safa_lp/
├── core/                           # 1. CAMADA FUNCIONAL (Lógica de Negócio)
│   ├── matriz_loader.py            # Leitura defensiva e cascata de matrizes JSON
│   ├── prompt_engine.py            # Mecanismo de injeção de placeholders nos templates
│   ├── validator.py                # Validador estrutural e consistência gabarito/explicação
│   ├── item_parser.py              # Parser Regex de lote multiline (^#\s*Item\s+\d+)
│   ├── word_formatter.py           # Construtor de arquivos DOCX individuais
│   ├── safa_processor.py           # Segmentador de paginação estrita (6 páginas) SAFA
│   ├── png_converter.py            # Automação win32com, ponte PDF e renderização Pillow
│   └── logger.py                   # Central de auditoria e persistência de telemetria
├── data/                           # 2. CAMADA DE DADOS EDUCACIONAIS
│   └── matrizes_portugues.json     # Árvore de descritores, anos e estados de ativação
├── templates/                      # 3. CAMADA DE INJEÇÃO PEDAGÓGICA (Templates de Prompt)
│   ├── prompt_portugues_novo_item.txt
│   └── prompt_portugues_revisao_item.txt
├── layout_interface/               # 4. CAMADA DE INTERFACE DO USUÁRIO (Aparência da Tela)
│   ├── index.html                  # Estrutura semântica das 4 abas operacionais
│   ├── style.css                   # Definição de cores, grids, estados de hover e validações
│   └── scripts.js                  # Comportamento de tela exclusivo (Tab switching e Clipboard)
├── layout_documento/               # 5. CAMADA DOCUMENTAL (Estilos de Impressão/Word)
│   └── layout_safa.json            # Metadados de geometria da página (fontes, margens, caixas)
├── logs/                           # 6. ARTEFATOS DE AUDITORIA
│   └── relatorio_processamento.json# Histórico central cumulativo de execuções
├── tests/                          # 7. CAMADA DE HOMOLOGAÇÃO E QUALIDADE
│   ├── test_suite_safa.py          # Bateria integrada de asserções lógicas
│   ├── gerar_relatorio.py          # Script de automação e compilação do relatório
│   └── RELATORIO_TESTES.md         # Sumário descritivo dos status de homologação
├── version.json                    # Controle de build e tags do sistema
├── requirements.txt                # Manifesto de bibliotecas externas
└── main.py                         # Ponto de entrada (Bootstrap da aplicação)
```

---

## 4. Funcionamento Detalhado das Quatro Abas

### 1. Formulador de Prompt
* **Objetivo**: Guiar o usuário na estruturação de prompts assertivos para LLMs.
* **Mecanismo**: Lê reativamente `data/matrizes_portugues.json` populando de forma encadeada os campos de seleção. O motor injeta os dados nos templates.
* **Regra de Negócio Crítica**: Executa validação de adaptação controlada se o Ano-Alvo divergir da Etapa da matriz.

### 2. Formatador de Item
* **Objetivo**: Validar se o item gerado possui consistência sintática mínima.
* **Mecanismo**: Analisa o campo de texto bruto em busca de marcadores de seção imperativos.

### 3. Padrão SAFA
* **Objetivo**: Fatiar cadernos de itens em lote e aplicar a diagramação padrão de homologação.
* **Mecanismo**: Itens são convertidos em arquivos de exatamente 6 páginas. Cada alternativa é isolada em sua respectiva página.

### 4. Montagem das Imagens
* **Objetivo**: Gerar ativos visuais bitmap estáticos (.png) em 200 DPI via automação Word COM.

---

## 5. Guias de Customização e Manutenção
(Ver arquivo original para detalhes sobre atualização de matrizes, templates e layouts)

---

## 6. Interpretação de Logs e Diagnóstico de Falhas
Toda transação operacional registra uma entrada consolidada em `logs/relatorio_processamento.json`.

---

## 7. Fluxo de Uso Recomendado
1. **Configuração Inicial**: Valide o arquivo `version.json`.
2. **Fase de Formulação**: Utilize a aba **Formulador de Prompt**.
3. **Fase de Importação**: Insira a marcação `# Item X` no topo de cada item.
4. **Fase de Diagramação**: Utilize a aba **Padrão SAFA**.
5. **Fase de Extração**: Utilize a aba **Montagem das Imagens**.
