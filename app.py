import streamlit as st
import pdfplumber
import fitz  # PyMuPDF para extrair imagens
import pandas as pd
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image as ExcelImage

# 1. Configuração da Página
st.set_page_config(page_title="Gerador de Relatório Fotográfico", layout="wide", page_icon="📸")

st.title("📸 Gerador de Relatório Fotográfico Automático")
st.markdown("Extraia os dados dos PDFs base e combine automaticamente com as fotos do TRO.")

# 2. Interface de Upload e Texto
col1, col2 = st.columns(2)

with col1:
    uploader_dados = st.file_uploader("1. PDFs de Dados (Base)", type=['pdf'], accept_multiple_files=True)
    caixa_texto = st.text_area("2. Ou cole o texto base aqui...", height=150)

with col2:
    uploader_tro = st.file_uploader("3. Opcional: PDF TRO (Fotos)", type=['pdf'], accept_multiple_files=False)

# 3. Padrão Regex (Mantido igual)
padrao_base = r"([A-Z]{2}-\d{3}-[A-Z]{2})\s+(.*?)\s+(Sentido\s+(?:Crescente|Decrescente))\s*-\s*([\d,]+)\s+([\d,]+)"

def extrair_padrao(texto_cru):
    dados = []
    if texto_cru:
        texto_limpo = re.sub(r'\s+', ' ', texto_cru)
        encontros = re.finditer(padrao_base, texto_limpo)
        for match in encontros:
            km_inicial = match.group(4)
            texto_formatado = f"{match.group(1)} {match.group(2).strip()} {match.group(3)} - {km_inicial} {match.group(5)}"
            # Guarda o KM inicial limpo (ex: "758.700" ou "758,70") para facilitar a busca
            km_chave = km_inicial.replace(',', '.') 
            dados.append({"texto": texto_formatado, "km_chave": km_chave})
    return dados

# 4. Função para extrair imagens do TRO
def extrair_imagens_tro(conteudo_pdf_bytes):
    mapa_imagens = {}
    doc = fitz.open(stream=conteudo_pdf_bytes, filetype="pdf")
    
    for pagina in doc:
        lista_imagens = pagina.get_images(full=True)
        dicionario_texto = pagina.get_text("dict")
        
        for img_info in lista_imagens:
            xref = img_info[0]
            imagem_bytes = doc.extract_image(xref)["image"]
            
            blocos = dicionario_texto.get("blocks", [])
            km_encontrado = None
            
            for bloco in blocos:
                if "lines" in bloco:
                    for linha in bloco["lines"]:
                        for span in linha["spans"]:
                            texto = span["text"].strip()
                            # Tenta identificar número com vírgula (ex: 758,70)
                            if re.match(r"^\d{1,4},\d{1,3}$", texto): 
                                km_encontrado = texto.replace(',', '.')
                                break
                        if km_encontrado: break
                if km_encontrado: break
            
            if km_encontrado:
                mapa_imagens[km_encontrado] = imagem_bytes
                
    return mapa_imagens

# 5. Processamento e Criação do Excel
st.divider()

if st.button("4. Gerar Planilha Completa", type="primary", use_container_width=True):
    with st.spinner("Iniciando a automação... Processando textos e procurando fotos..."):
        
        registros = []
        mapa_fotos = {}
        
        # A. Processar Texto Base
        if caixa_texto.strip():
            registros.extend(extrair_padrao(caixa_texto.strip()))
            
        # B. Processar PDFs Base
        if uploader_dados:
            for arquivo in uploader_dados:
                # O Streamlit devolve os ficheiros como objetos que precisam de ser lidos
                with pdfplumber.open(io.BytesIO(arquivo.read())) as pdf:
                    for pagina in pdf.pages:
                        texto_pagina = pagina.extract_text()
                        if texto_pagina:
                            registros.extend(extrair_padrao(texto_pagina))
        
        # C. Processar o PDF com as Fotos (TRO)
        if uploader_tro:
            mapa_fotos = extrair_imagens_tro(uploader_tro.read())
            st.success(f"📸 Fantástico! Foram encontradas e recortadas {len(mapa_fotos)} fotos no arquivo TRO.")

        # 6. Construir o Excel
        if registros:
            wb = Workbook()
            ws = wb.active
            ws.title = "Relatório Fotográfico"
            
            ws.column_dimensions['A'].width = 5
            ws.column_dimensions['B'].width = 45
            ws.column_dimensions['C'].width = 45
            ws.column_dimensions['D'].width = 5
            
            linha_excel = 1 
            fotos_inseridas = 0
            
            for reg in registros:
                texto_final = reg['texto']
                km_buscado = reg['km_chave']
                
                # --- LINHA DA FOTO ---
                ws.row_dimensions[linha_excel].height = 150
                
                foto_bytes = None
                for km_tro, img_data in mapa_fotos.items():
                    if km_tro in km_buscado or km_buscado.startswith(km_tro):
                        foto_bytes = img_data
                        break
                
                if foto_bytes:
                    imagem_io = io.BytesIO(foto_bytes)
                    img_excel = ExcelImage(imagem_io)
                    img_excel.height = 180 
                    img_excel.width = 300
                    
                    posicao = f"B{linha_excel}"
                    ws.add_image(img_excel, posicao)
                    fotos_inseridas += 1
                else:
                    celula_b = ws.cell(row=linha_excel, column=2)
                    celula_b.value = "FOTO NÃO ENCONTRADA"
                    celula_b.alignment = Alignment(horizontal='center', vertical='center')

                # --- LINHA DO TEXTO ---
                linha_texto_excel = linha_excel + 1
                ws.row_dimensions[linha_texto_excel].height = 11.25
                ws.merge_cells(start_row=linha_texto_excel, start_column=2, end_row=linha_texto_excel, end_column=3)
                celula_texto = ws.cell(row=linha_texto_excel, column=2)
                celula_texto.value = texto_final
                celula_texto.alignment = Alignment(horizontal='center', vertical='center')
                
                linha_excel += 2
            
            # 7. Disponibilizar o Botão de Download
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            
            st.success(f"✅ Feito! **{len(registros)}** linhas criadas e **{fotos_inseridas}** fotos coladas com sucesso.")
            
            st.download_button(
                label="📥 Baixar Planilha Final (Com Fotos)",
                data=excel_buffer,
                file_name="Relatorio_Com_Fotos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
        else:
            st.error("⚠️ Nenhum dado de rodovia foi encontrado no texto ou nos PDFs base fornecidos.")
