import streamlit as st
import pdfplumber
import fitz  # PyMuPDF para extrair imagens de PDFs
import pandas as pd
import re
import io
import ipywidgets as widgets
from IPython.display import display, clear_output, HTML
import base64
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage
import pytesseract # Motor de leitura de imagens (OCR)

# 1. Interface Atualizada
caixa_texto = widgets.Textarea(placeholder='Cole o texto base de rodovias e KMs aqui...', layout=widgets.Layout(width='100%', height='100px'))
uploader_antes = widgets.FileUpload(accept='.pdf, .jpg, .jpeg, .png', multiple=True, description='Inserir fotos (ANTES)', button_style='info', layout=widgets.Layout(width='max-content'))
uploader_depois = widgets.FileUpload(accept='.pdf, .jpg, .jpeg, .png', multiple=True, description='Inserir fotos (DEPOIS', button_style='warning', layout=widgets.Layout(width='max-content'))

btn_processar = widgets.Button(description='Gerar Planilha Completa', button_style='success', icon='file-excel-o')
btn_limpar = widgets.Button(description='Limpar Tudo', button_style='danger', icon='trash')
out = widgets.Output()

# 2. Padrão Regex para o texto base
padrao_base = r"([A-Z]{2}-\d{3}-[A-Z]{2})\s+(.*?)\s+(Sentido\s+(?:Crescente|Decrescente))\s*-\s*([\d,]+)\s+([\d,]+)"
padrao_km_solto = r"(\d{1,4}[.,]\d{1,3})" # Padrão para caçar KM solto em títulos ou imagens

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

# Função poderosa para ler Título, PDF e Imagem
def mapear_fotos(arquivos_upload):
    mapa_arquivos = {}

    # Tratamento para compatibilidade de versões do ipywidgets
    lista_arquivos = arquivos_upload if isinstance(arquivos_upload, (tuple, list)) else arquivos_upload.values()

    for arquivo in lista_arquivos:
        nome_arquivo = arquivo['metadata']['name'] if 'metadata' in arquivo else arquivo['name']
        conteudo = arquivo['content']
        extensao = nome_arquivo.split('.')[-1].lower()

        # 1. Tenta achar o KM no título do arquivo (Ex: "Foto_Antes_KM_758,70.jpg")
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
                    pass # Se falhar o OCR, segue a vida

            if km_encontrado:
                mapa_arquivos[km_encontrado] = conteudo

    return mapa_arquivos

# 3. Processamento Principal
def acao_processar(b):
    with out:
        clear_output()
        print("🔍 Analisando dados base e varrendo arquivos em busca de KMs correspondentes...")

        registros = extrair_padrao_texto(caixa_texto.value.strip())

        # Extrai imagens dos uploads e cria os "dicionários" de busca
        mapa_antes = mapear_fotos(uploader_antes.value) if uploader_antes.value else {}
        mapa_depois = mapear_fotos(uploader_depois.value) if uploader_depois.value else {}

        print(f"✅ Encontrados {len(mapa_antes)} KMs nas fotos 'ANTES' e {len(mapa_depois)} KMs nas fotos 'DEPOIS'.")

        if registros:
            wb = Workbook()
            ws = wb.active
            ws.title = "Relatório Fotográfico"

            ws.column_dimensions['A'].width = 5
            ws.column_dimensions['B'].width = 45 # Antes
            ws.column_dimensions['C'].width = 45 # Depois
            ws.column_dimensions['D'].width = 5

            # --- CABEÇALHO DO LAYOUT ---
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

                # --- LINHA DA FOTO (Altura 150px) ---
                ws.row_dimensions[linha_excel].height = 150

                # Lógica para achar a foto certa: ignora zeros à direita (Ex: 758.700 -> 758.7)
                foto_antes_bytes = next((img for km, img in mapa_antes.items() if km in km_buscado or km_buscado.startswith(km)), None)
                foto_depois_bytes = next((img for km, img in mapa_depois.items() if km in km_buscado or km_buscado.startswith(km)), None)

                # Colar Foto Antes
                if foto_antes_bytes:
                    img_excel = ExcelImage(io.BytesIO(foto_antes_bytes))
                    img_excel.height, img_excel.width = 180, 300
                    ws.add_image(img_excel, f"B{linha_excel}")
                    fotos_coladas += 1
                else:
                    ws[f"B{linha_excel}"] = "[ SEM FOTO ANTES ]"
                    ws[f"B{linha_excel}"].alignment = Alignment(horizontal='center', vertical='center')

                # Colar Foto Depois
                if foto_depois_bytes:
                    img_excel = ExcelImage(io.BytesIO(foto_depois_bytes))
                    img_excel.height, img_excel.width = 180, 300
                    ws.add_image(img_excel, f"C{linha_excel}")
                    fotos_coladas += 1
                else:
                    ws[f"C{linha_excel}"] = "[ SEM FOTO DEPOIS ]"
                    ws[f"C{linha_excel}"].alignment = Alignment(horizontal='center', vertical='center')

                # --- LINHA DO TEXTO (Altura 15px, Mesclada) ---
                linha_texto_excel = linha_excel + 1
                ws.row_dimensions[linha_texto_excel].height = 11.25
                ws.merge_cells(start_row=linha_texto_excel, start_column=2, end_row=linha_texto_excel, end_column=3)
                ws.cell(row=linha_texto_excel, column=2, value=texto_final).alignment = Alignment(horizontal='center', vertical='center')

                linha_excel += 2

            # Finalizar e gerar Download
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            b64 = base64.b64encode(excel_buffer.getvalue()).decode()

            html_download = f'''
            <div style="margin-top: 20px;">
                <p>✅ <b>{len(registros)}</b> blocos criados | 📸 <b>{fotos_coladas}</b> fotos correspondentes encaixadas!</p>
                <a download="Relatorio_Antes_Depois.xlsx" href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}">
                    <button style="background-color: #28a745; color: white; padding: 12px 24px; border: none; border-radius: 5px; cursor: pointer; font-weight: bold;">
                        📥 Baixar Relatório Final
                    </button>
                </a>
            </div>
            '''
            display(HTML(html_download))
        else:
            print("⚠️ Nenhum KM válido foi encontrado no texto base.")

def acao_limpar(b):
    with out:
        clear_output()
    caixa_texto.value = ''
    for up in [uploader_antes, uploader_depois]:
        if isinstance(up.value, tuple): up.value = ()
        else: up.value.clear(); up._counter = 0
    print("✨ Ambiente limpo!")

btn_processar.on_click(acao_processar)
btn_limpar.on_click(acao_limpar)

display(widgets.VBox([
    caixa_texto,
    widgets.HBox([uploader_antes, uploader_depois]),
    widgets.HBox([btn_processar, btn_limpar]),
    out
]))
