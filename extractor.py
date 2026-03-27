"""
extractor.py
Funções para extração de texto de imagens de páginas de jornal.

Fluxo: imagem JPG → OCR (pytesseract) → API de visão (OpenRouter/Grok) → Excel

Formato esperado dos arquivos de imagem: YYYY-MM-DD_PAGINA.jpg
Exemplo: 2012-06-04_002.jpg
"""

import os
import re
import base64
import json
import pandas as pd
from PIL import Image
import pytesseract


# ─── Funções auxiliares ───────────────────────────────────────────────────────

def encode_image(image_path):
    """Converte uma imagem para base64 (necessário para enviar à API)."""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")


def ocr_from_image(image_path):
    """Extrai texto bruto de uma imagem usando OCR (pytesseract)."""
    try:
        image = Image.open(image_path)
        return pytesseract.image_to_string(image)
    except Exception:
        return "Não foi possível extrair o texto via OCR, use apenas a imagem."


def extract_date_from_filename(filename):
    """
    Extrai a data e o ano de um nome de arquivo no formato YYYY-MM-DD_PAGINA.jpg.
    Retorna (data, ano). Se não encontrar, retorna ("Desconhecido", "Desconhecido").
    """
    match = re.search(r"(\d{4}-\d{2}-\d{2})", filename)
    if match:
        date = match.group(1)
        year = date.split("-")[0]
        return date, year
    return "Desconhecido", "Desconhecido"


def extract_page_from_filename(filename):
    """
    Extrai o número de página de um arquivo no formato YYYY-MM-DD_PAGINA.jpg.
    Retorna o número como inteiro, ou None se não encontrar.
    """
    try:
        page_part = filename.split("_")[1]
        return int(re.sub(r"\D", "", page_part))
    except Exception:
        return None


# ─── Prompt padrão ────────────────────────────────────────────────────────────

DEFAULT_PROMPT = """Você receberá uma imagem de uma página de jornal e seu texto extraído por OCR (que pode conter erros).
Extraia o texto corretamente, corrigindo os erros do OCR a partir da imagem, e organize o resultado no formato pedido abaixo.

INSTRUÇÕES:
- Extraia os títulos de cada matéria (estão em fonte maior e em negrito).
- Para cada título, extraia o autor (normalmente logo acima ou abaixo do título). Se não houver autor, use "Desconhecido".
- Para cada título, extraia TODO o texto associado, incluindo todos os parágrafos e colunas — o texto pode ser longo.
- Seja fiel ao texto original, não invente nada.
- Ignore: títulos de seção (ex: "Esportes", "Cultura"), anúncios, legendas de fotos, cartas de leitores.

Responda APENAS com um JSON neste formato (sem texto adicional):
[
    {{"título": "Título do Artigo 1", "autor": "Nome do Autor", "texto": "Texto completo do artigo..."}},
    {{"título": "Título do Artigo 2", "autor": "Desconhecido", "texto": "Texto completo do artigo..."}}
]

TEXTO OCR:
{ocr_text}

IMAGEM DO JORNAL:"""


# ─── Função principal de extração ─────────────────────────────────────────────

def extract_articles_from_image(image_path, client, model="x-ai/grok-4-fast", prompt_template=None):
    """
    Extrai artigos de uma imagem de jornal combinando OCR e API de visão.

    Args:
        image_path (str): Caminho para a imagem (.jpg/.jpeg).
        client (OpenAI): Cliente da API (OpenRouter ou compatível).
        model (str): Modelo de visão a usar.
        prompt_template (str): Template do prompt. Use {ocr_text} como placeholder.
                               Se None, usa o prompt padrão.

    Returns:
        list[dict]: Lista de artigos com chaves 'título', 'autor', 'texto'.
    """
    if prompt_template is None:
        prompt_template = DEFAULT_PROMPT

    ocr_text = ocr_from_image(image_path)
    base64_image = encode_image(image_path)
    prompt = prompt_template.format(ocr_text=ocr_text)

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "system",
                    "content": "Você é um assistente especializado em extrair textos de imagens de jornais.",
                },
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"},
                        },
                    ],
                },
            ],
        )
        raw = response.choices[0].message.content.strip()
        return json.loads(raw)

    except json.JSONDecodeError:
        print(f"  Aviso: resposta não era JSON válido para {os.path.basename(image_path)}")
        return [{"título": "Erro na extração", "autor": "Erro na extração", "texto": "Erro na extração"}]

    except Exception as e:
        print(f"  Erro ao processar {os.path.basename(image_path)}: {e}")
        return [{"título": "Erro na extração", "autor": "Erro na extração", "texto": str(e)}]


# ─── Processamento em lote ────────────────────────────────────────────────────

def process_folder(input_folder, output_xlsx, client, news_name, model="x-ai/grok-4-fast", save_every=20, prompt_template=None):
    """
    Processa todas as imagens de uma pasta e salva os resultados em um arquivo Excel.

    Se o arquivo de saída já existir, retoma de onde parou (pulando imagens já processadas).

    Args:
        input_folder (str): Pasta com as imagens (.jpg/.jpeg).
        output_xlsx (str): Caminho do arquivo Excel de saída.
        client (OpenAI): Cliente da API (OpenRouter ou compatível).
        news_name (str): Nome do jornal para identificação (ex: "FSP", "GLO", "ESP").
        model (str): Modelo de visão a usar.
        save_every (int): Salvar o arquivo a cada N imagens processadas.
        prompt_template (str): Template do prompt. Se None, usa o prompt padrão.

    Returns:
        pd.DataFrame: Tabela com todos os artigos extraídos.
    """
    columns = ["Jornal", "Ano", "Arquivo", "Data", "Página", "Título", "Autor", "Texto Associado"]
    results = []
    processed_files = set()

    # Retomar de onde parou, se o arquivo de saída já existir
    if os.path.exists(output_xlsx):
        df_existing = pd.read_excel(output_xlsx)
        df_existing = df_existing[df_existing["Título"] != "Erro na extração"]
        results = df_existing.values.tolist()
        processed_files = set(df_existing["Arquivo"].tolist())
        print(f"Retomando: {len(processed_files)} arquivos já processados anteriormente.")

    # Listar imagens na pasta
    all_images = []
    for root, _, files in os.walk(input_folder):
        for filename in sorted(files):
            if filename.lower().endswith((".jpg", ".jpeg")):
                all_images.append((root, filename))

    total = len(all_images)
    pending_count = sum(1 for _, f in all_images if f not in processed_files)
    print(f"Total de imagens encontradas: {total} | A processar: {pending_count}")

    processed_now = 0
    for root, filename in all_images:
        if filename in processed_files:
            continue

        image_path = os.path.join(root, filename)
        date, year = extract_date_from_filename(filename)
        page = extract_page_from_filename(filename)

        articles = extract_articles_from_image(image_path, client, model=model, prompt_template=prompt_template)

        for article in articles:
            results.append([
                news_name,
                year,
                filename,
                date,
                page,
                article.get("título", "Desconhecido"),
                article.get("autor", "Desconhecido"),
                article.get("texto", "Desconhecido"),
            ])

        processed_now += 1
        print(f"[{processed_now}/{pending_count}] {filename} → {len(articles)} artigo(s) extraído(s)")

        if processed_now % save_every == 0:
            df = pd.DataFrame(results, columns=columns)
            df.to_excel(output_xlsx, index=False)
            print(f"  Salvo intermediário após {processed_now} arquivos.")

    # Salvar resultado final
    df = pd.DataFrame(results, columns=columns)
    df.to_excel(output_xlsx, index=False)
    print(f"\nConcluído! {processed_now} imagens processadas.")
    print(f"Resultado salvo em: {output_xlsx}")
    return df
