# Extração de Textos de Imagens de Jornais

## Autores
![Ana Carolina Erthal](https://img.shields.io/badge/Ana%20Carolina%20Erthal-FGV--EMAp-blue)
![João Feres Júnior](https://img.shields.io/badge/João%20Feres%20Júnior-IESP--UERJ-blue)


Ferramenta para extrair automaticamente os textos de imagens de páginas de jornal, organizando os resultados em um arquivo Excel.

> 🔗 **[Acesse a página do projeto](https://iesp-uerj.github.io/lemep_newspaper_extract/)**


---

## Como funciona

Para cada imagem de página de jornal, o processo é:

1. **OCR** (pytesseract) faz uma leitura inicial do texto, sem estrutra definida e com possíveis erros de reconhecimento e ordenação.
2. **API de visão** (OpenRouter + LLM) recebe a imagem e o texto do OCR, corrige os erros e extrai os artigos de forma estruturada
3. Os resultados são salvos em um **arquivo Excel** com título, autor e texto completo de cada matéria

```
imagem JPG  →  OCR  →  API de visão  →  Excel
```

---

## Instalação

### 1. Dependências Python

```bash
pip install openai pandas pillow pytesseract openpyxl
```

### 2. Tesseract OCR

O Tesseract precisa ser instalado separadamente no sistema:

- **Mac:** `brew install tesseract tesseract-lang`
- **Linux:** `sudo apt install tesseract-ocr tesseract-ocr-por`
- **Windows:** baixar o instalador em https://github.com/UB-Mannheim/tesseract/wiki

### 3. Chave de API

Crie uma conta no [OpenRouter](https://openrouter.ai) e obtenha sua chave de API.

---

## Uso

### Formato esperado das imagens

Os arquivos devem seguir o padrão: `YYYY-MM-DD_PAGINA.jpg`

```
2012-06-04_002.jpg   →  edição de 4 de junho de 2012, página 2
2010-11-15_003.jpg   →  edição de 15 de novembro de 2010, página 3
```

### Pelo notebook (recomendado)

Abra o arquivo `extractor.ipynb` e preencha as 4 variáveis na célula de configuração:

```python
API_KEY      = "sua-chave-aqui"
INPUT_FOLDER = "caminho/para/as/imagens"
OUTPUT_XLSX  = "resultados.xlsx"
JORNAL_NAME  = "FSP"
```

Depois, execute as células em ordem.

### Pelo Python

```python
from openai import OpenAI
from extractor import process_folder

client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key="sua-chave-aqui",
)

df = process_folder(
    input_folder="caminho/para/as/imagens",
    output_xlsx="resultados.xlsx",
    client=client,
    news_name="FSP",
)
```

---

## Resultado

O arquivo Excel gerado contém uma linha por artigo extraído:

| Coluna | Descrição |
|---|---|
| Jornal | Nome do jornal informado |
| Ano | Ano da edição (extraído do nome do arquivo) |
| Arquivo | Nome do arquivo de imagem original |
| Data | Data da edição (extraída do nome do arquivo) |
| Página | Número da página |
| Título | Título do artigo |
| Autor | Autor do artigo (ou "Desconhecido") |
| Texto Associado | Texto completo do artigo |

---

## Retomada automática

Se a execução for interrompida, basta rodar novamente. O script detecta automaticamente quais imagens já foram processadas e continua de onde parou.

---

## Detecção de páginas faltantes

O notebook `detect_missing_pages.ipynb` verifica a integridade da coleção de imagens antes de rodar a extração.

Ele detecta dois tipos de problema:
- **Datas parcialmente incompletas**: existem arquivos, mas faltam páginas específicas
- **Datas completamente ausentes**: nenhum arquivo naquele dia (requer definir `START_DATE` e `END_DATE`)

```python
INPUT_FOLDER   = "caminho/para/as/imagens"
OUTPUT_TXT     = "missing_pages_report.txt"
REQUIRED_PAGES = [2, 3]    # ou None
START_DATE     = "2003-01-01"  # ou None
END_DATE       = "2013-12-31"  # ou None
```

O resultado é um arquivo `.txt` com um resumo e a lista detalhada por data.

---

## Estrutura dos arquivos

```
lemep_newspaper_extract/
├── extractor.py                  # Funções principais de extração
├── extractor.ipynb               # Notebook para extração de texto
├── detect_missing_pages.ipynb    # Notebook para verificar integridade da coleção
├── index.html                    # Página do projeto
├── data/                         # Dados de saída
└── README.md                     # Este arquivo
```
