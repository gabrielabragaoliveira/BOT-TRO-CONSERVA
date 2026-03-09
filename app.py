import streamlit as st
import pdfplumber
import fitz  # PyMuPDF para extrair imagens de PDFs
import pandas as pd
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage
import pytesseract # Motor de leitura de imagens (OCR)

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador de Relatório Fotográfico", page_icon="📸", layout="wide")

st.title("📸 Gerador de Relatório: Antes e Depois")
st.markdown("""
Esta ferramenta cruza os KMs do seu texto base com as fotos fornecidas (via título, texto do PDF ou leitura OCR da imagem) 
e gera uma planilha Excel formatada automaticamente.
""")

# --- 2. PADRÕES REGEX ---
padrao_base = r"([A-Z]{2}-\d{3}-[A-Z]{2})\s+(.*?)\s+(Sentido\s+(?:Crescente|Decrescente))\s*-\s*([\d,]+)\s+([\d,]+)"
padrao_km_solto = r"(\d{1,4}[.,]\d{1,3})" 

def extrair_padrao_texto(texto_cru):
    dados = []
    if texto_cru:
        texto_limpo = re.sub(r'\s+', ' ', texto_cru)
        encontros = re.finditer(padrao_base, texto_limpo)
        for match in encontros:
            km_inicial = match.group(4)
            texto_formatado = f"{match.group(1)} {match.group(2).strip()} {match.group(3)} - {km_inicial} {match.group(5)}"
            km_chave = km_inicial.replace(',', '.')
            dados.append({"texto": texto_formatado, "km_chave": km_chave})
    return dados

# --- 3. FUNÇÃO DE MAPEAMENTO DE FOTOS ---
def mapear_fotos(arquivos_upload):
    mapa_arquivos = {}

    for arquivo in arquivos_upload:
        nome_arquivo = arquivo.name
        conteudo = arquivo.read()
        extensao = nome_arquivo.split('.')[-1].lower()

        # 1. Tenta achar o KM no título do arquivo
        km_no_titulo = re.search(padrao_km_solto, nome_arquivo)

        # Se for PDF
        if extensao == 'pdf':
            doc = fitz.open(stream=conteudo, filetype="pdf")
            for pagina in doc:
                texto_pdf = pagina.get_text()
                km_encontrado = None

                # 2. Procura no texto do PDF
                match_texto = re.search(padrao_km_solto, texto_pdf)
                if match_texto:
                    km_encontrado = match_texto.group(1).replace(',', '.')
                elif km_no_titulo:
                    km_encontrado = km_no_titulo.group(1).replace(',', '.')

                # Extrai a maior imagem da página para ligar ao KM
                imagens = pagina.get_images(full=True)
                if imagens and km_encontrado:
                    xref = imagens[0][0]
                    imagem_bytes = doc.extract_image(xref)["image"]
                    mapa_arquivos[km_encontrado] = imagem_bytes

        # Se for Imagem Direta (JPG, PNG)
        elif extensao in ['jpg', 'jpeg', 'png']:
            km_encontrado = None
            if km_no_titulo:
                km_encontrado = km_no_titulo.group(1).replace(',', '.')
            else:
                # 3. OCR: Lê o texto carimbado dentro da imagem
                try:
                    img_pil = PILImage.open(io.BytesIO(conteudo))
                    texto_imagem = pytesseract.image_to_string(img_pil)
                    match_ocr = re.search(padrao_km_solto, texto_imagem)
                    if match_ocr:
                        km_encontrado = match_ocr.group(1).replace(',', '.')
                except Exception as e:
                    pass # Ignora erros de OCR silenciosamente

            if km_encontrado:
                mapa_arquivos[km_encontrado] = conteudo

    return mapa_arquivos

# --- 4. INTERFACE STREAMLIT ---
texto_base = st.text_area("📝 Texto Base (Cole as rodovias e KMs aqui):", height=150)

col1, col2 = st.columns(2)
with col1:
    arquivos_antes = st.file_uploader("📥 Inserir fotos (ANTES)", accept_multiple_files=True, type=['pdf', 'jpg', 'jpeg', 'png'])
with col2:
    arquivos_depois = st.file_uploader("📥 Inserir fotos (DEPOIS)", accept_multiple_files=True, type=['pdf', 'jpg', 'jpeg', 'png'])

# --- 5. LÓGICA DE PROCESSAMENTO ---
if st.button("🚀 Gerar Planilha Completa", type="primary", use_container_width=True):
    if not texto_base.strip():
        st.error("⚠️ Por favor, cole o texto base de rodovias para começarmos.")
    else:
        with st.spinner("Analisando KMs e cruzando com as fotos. Isso pode levar alguns segundos..."):
            registros = extrair_padrao_texto(texto_base.strip())
            
            mapa_antes = mapear_fotos(arquivos_antes) if arquivos_antes else {}
            mapa_depois = mapear_fotos(arquivos_depois) if arquivos_depois else {}

            st.success(f"✅ Encontrados {len(mapa_antes)} KMs nas fotos 'ANTES' e {len(mapa_depois)} KMs nas fotos 'DEPOIS'.")

            if registros:
                wb = Workbook()
                ws = wb.active
                ws.title = "Relatório Fotográfico"

                # Configuração das Colunas
                ws.column_dimensions['A'].width = 5
                ws.column_dimensions['B'].width = 45 # Antes
                ws.column_dimensions['C'].width = 45 # Depois
                ws.column_dimensions['D'].width = 5

                # Cabeçalho Fixo
                ws.merge_cells('B1:C1')
                ws['B1'] = "RELATÓRIO DE EXECUÇÃO"
                ws['B1'].alignment = Alignment(horizontal='center', vertical='center')
                ws['B1'].font = Font(bold=True, color="FFFFFF")
                ws['B1'].fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")

                ws['B2'] = "Foto(s) do Local Antes / Durante a Execução"
                ws['C2'] = "Foto(s) do Local Após a Execução"
                for celula in ['B2', 'C2']:
                    ws[celula].alignment = Alignment(horizontal='center', vertical='center')
                    ws[celula].font = Font(bold=True)
                    ws[celula].fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

                linha_excel = 3
                fotos_coladas = 0

                for reg in registros:
                    texto_final = reg['texto']
                    km_buscado = reg['km_chave']

                    # Linha da Foto
                    ws.row_dimensions[linha_excel].height = 150

                    # Busca foto Antes e Depois
                    foto_antes_bytes = next((img for km, img in mapa_antes.items() if km in km_buscado or km_buscado.startswith(km)), None)
                    foto_depois_bytes = next((img for km, img in mapa_depois.items() if km in km_buscado or km_buscado.startswith(km)), None)

                    # Cola Foto Antes
                    if foto_antes_bytes:
                        img_excel = ExcelImage(io.BytesIO(foto_antes_bytes))
                        img_excel.height, img_excel.width = 180, 300
                        ws.add_image(img_excel, f"B{linha_excel}")
                        fotos_coladas += 1
                    else:
                        ws[f"B{linha_excel}"] = "[ SEM FOTO ANTES ]"
                        ws[f"B{linha_excel}"].alignment = Alignment(horizontal='center', vertical='center')

                    # Cola Foto Depois
                    if foto_depois_bytes:
                        img_excel = ExcelImage(io.BytesIO(foto_depois_bytes))
                        img_excel.height, img_excel.width = 180, 300
                        ws.add_image(img_excel, f"C{linha_excel}")
                        fotos_coladas += 1
                    else:
                        ws[f"C{linha_excel}"] = "[ SEM FOTO DEPOIS ]"
                        ws[f"C{linha_excel}"].alignment = Alignment(horizontal='center', vertical='center')

                    # Linha do Texto (Legenda)
                    linha_texto_excel = linha_excel + 1
                    ws.row_dimensions[linha_texto_excel].height = 11.25
                    ws.merge_cells(start_row=linha_texto_excel, start_column=2, end_row=linha_texto_excel, end_column=3)
                    ws.cell(row=linha_texto_excel, column=2, value=texto_final).alignment = Alignment(horizontal='center', vertical='center')

                    linha_excel += 2

                # Prepara o arquivo para o botão de download
                excel_buffer = io.BytesIO()
                wb.save(excel_buffer)
                
                st.info(f"📊 Estrutura pronta! {len(registros)} registros processados com {fotos_coladas} fotos inseridas.")

                st.download_button(
                    label="📥 Clique para Baixar o Relatório (Excel)",
                    data=excel_buffer.getvalue(),
                    file_name="Relatorio_Fotos_Antes_Depois.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            else:
                st.warning("⚠️ Nenhum padrão válido de Rodovia/KM foi encontrado no texto.")
