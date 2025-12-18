import streamlit as st
import pandas as pd
import io
import os
import tempfile
from PIL import Image

# --- IMPORTS PARA DOCX ---
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx2pdf import convert

# --- IMPORTS PARA PDF ---
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA E CSS
# ==========================================
st.set_page_config(page_title="Gerador de Relatórios DP", layout="wide", page_icon="⚓")

# CSS para visual corporativo (Azul Bram/Edison Chouest)
st.markdown("""
    <style>
    :root { --primary-color: #0054a6; }
    div.stButton > button:first-child {
        background-color: #0054a6;
        color: white;
        border-radius: 5px;
        border: none;
        font-weight: bold;
    }
    div.stButton > button:hover { background-color: #003f7f; color: white; }
    h1, h2, h3 { color: #333; font-family: 'Arial', sans-serif; }
    .stSidebar { background-color: #f0f2f6; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 2. BARRA LATERAL (CONFIGURAÇÕES)
# ==========================================
st.sidebar.header("⚙️ Configurações do Documento")

# --- Configuração da Planilha ---
st.sidebar.subheader("Excel / Planilha")
sheet_input = st.sidebar.text_input(
    "Nome ou Índice da Aba (Sheet Name)", 
    value="0", 
    help="Digite '0' para a primeira aba, ou o nome exato (ex: 'Planilha1')."
)

# --- Configuração de Fontes ---
st.sidebar.subheader("Tipografia")
font_name_cfg = st.sidebar.selectbox("Fonte Principal", ["Raleway", "Arial", "Calibri", "Times New Roman"], index=0)
font_size_h1 = st.sidebar.number_input("Tamanho Título (H1)", value=20, step=1)
font_size_h2 = st.sidebar.number_input("Tamanho Subtítulo (H2)", value=14, step=1)
font_size_body = st.sidebar.number_input("Tamanho Corpo", value=11, step=1)

# --- Configuração de Margens ---
st.sidebar.subheader("Margens (Polegadas)")
col_m1, col_m2 = st.sidebar.columns(2)
margin_top = col_m1.number_input("Topo", value=1.0, step=0.1)
margin_bottom = col_m2.number_input("Rodapé", value=1.0, step=0.1)
margin_left = col_m1.number_input("Esquerda", value=0.5, step=0.1)
margin_right = col_m2.number_input("Direita", value=0.5, step=0.1)

# Agrupando configurações
CONFIG = {
    "sheet_target": sheet_input,
    "font_name": font_name_cfg,
    "h1": font_size_h1,
    "h2": font_size_h2,
    "body": font_size_body,
    "m_top": margin_top,
    "m_bottom": margin_bottom,
    "m_left": margin_left,
    "m_right": margin_right
}

# Define a logo do documento (usada no DOCX e PDF)
LOGO_DOC_PATH = "logo.png"

# ==========================================
# 3. FUNÇÕES DO GERADOR (DOCX)
# ==========================================

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for element in tcPr.xpath('./w:tcBorders'):
        tcPr.remove(element)
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ['top', 'left', 'bottom', 'right']:
        if edge in kwargs:
            edge_data = kwargs[edge]
            element = OxmlElement(f"w:{edge}")
            for key, value in edge_data.items():
                element.set(qn(f"w:{key}"), str(value))
            tcBorders.append(element)
    tcPr.append(tcBorders)

default_border_settings = {
    "top": {"sz": 4, "val": "single", "color": "000000", "space": "0"},
    "bottom": {"sz": 4, "val": "single", "color": "000000", "space": "0"},
    "left": {"sz": 0, "val": "nil", "color": "auto", "space": "0"},
    "right": {"sz": 0, "val": "nil", "color": "auto", "space": "0"}
}

def create_chapter_cover(doc, chapter_title, cfg):
    if len(doc.paragraphs) > 0:
        doc.add_page_break()
    for _ in range(15):
        p = doc.add_paragraph("")
        p.paragraph_format.line_spacing = 1
    
    para = doc.add_paragraph()
    para.paragraph_format.line_spacing = 1
    run = para.add_run(chapter_title)
    run.font.name = cfg['font_name']
    run.font.size = Pt(cfg['h1'])
    run.font.underline = True
    run.bold = True
    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    for _ in range(7):
        p = doc.add_paragraph("")
        p.paragraph_format.line_spacing = 1
    doc.add_page_break()

def create_bordered_section(doc, label, content, cfg, no_bottom_border=False, extra_space_top=Pt(3), extra_space_after_content=Pt(3)):
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    cell.text = ''
    
    p_title = cell.add_paragraph()
    p_title.paragraph_format.space_before = extra_space_top
    p_title.paragraph_format.space_after = Pt(0)
    run_title = p_title.add_run(f"{label}:")
    run_title.font.name = cfg['font_name']
    run_title.bold = True
    run_title.font.size = Pt(cfg['h2'] - 2)
    
    p_content = cell.add_paragraph()
    p_content.paragraph_format.space_before = Pt(0)
    p_content.paragraph_format.space_after = extra_space_after_content
    run_content = p_content.add_run(str(content))
    run_content.font.name = cfg['font_name']
    run_content.font.size = Pt(cfg['body'])
    
    if label.lower() == "method":
        p_content.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        cell.add_paragraph("")
        
    cell_border_settings = default_border_settings.copy()
    if no_bottom_border:
        cell_border_settings.pop("bottom", None)
    set_cell_border(cell, **cell_border_settings)

def create_test_page(doc, test_info, cfg):
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    
    para_title = doc.add_paragraph()
    para_title.paragraph_format.line_spacing = 1
    run_title = para_title.add_run(f"{test_info['Test']}")
    run_title.font.name = cfg['font_name']
    run_title.bold = True
    run_title.font.size = Pt(cfg['h2'])
    para_title.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    doc.add_paragraph("")
    
    create_bordered_section(doc, "Method", test_info['Method'], cfg, extra_space_top=Pt(1), extra_space_after_content=Pt(3))
    doc.add_paragraph("")
    
    para_steps_title = doc.add_paragraph()
    para_steps_title.paragraph_format.line_spacing = 1
    run_steps_title = para_steps_title.add_run("Steps:")
    run_steps_title.font.name = cfg['font_name']
    run_steps_title.bold = True
    run_steps_title.font.size = Pt(cfg['body'] + 1)
    
    for step in test_info['Steps']:
        p_step = doc.add_paragraph(f"{step}")
        p_step.style.font.name = cfg['font_name']
        p_step.style.font.size = Pt(cfg['body'])
        p_step.paragraph_format.line_spacing = 1
    doc.add_paragraph("")
    
    create_bordered_section(doc, "Expected Results", "\n".join(test_info['Expected Results']), cfg,
                             no_bottom_border=True, extra_space_top=Pt(1), extra_space_after_content=Pt(0))
    doc.add_paragraph("")
    
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    
    def format_cell(c, label, value):
        p = c.paragraphs[0]
        p.paragraph_format.line_spacing = 1
        r1 = p.add_run(label)
        r1.bold = True
        r1.font.name = cfg['font_name']
        r1.font.size = Pt(cfg['body'])
        if value:
            r2 = p.add_run(f" {value}")
            r2.font.name = cfg['font_name']
            r2.font.size = Pt(cfg['body'])
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    format_cell(table.cell(0, 0), "Max Deviation:", "")
    format_cell(table.cell(0, 1), "Position:", f"< {test_info['Max. Position Deviation (meters)']} meters")
    format_cell(table.cell(0, 2), "Heading:", f"< {test_info['Max. Heading Deviation (degrees)']} degrees")
    
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, **default_border_settings)
            
    doc.add_paragraph("")
    
    para_res_title = doc.add_paragraph()
    para_res_title.paragraph_format.line_spacing = 1
    run_res_title = para_res_title.add_run("Results:")
    run_res_title.font.name = cfg['font_name']
    run_res_title.bold = True
    run_res_title.font.size = Pt(cfg['body'] + 1)
    
    result_comments = test_info.get('Result + Comment')
    if result_comments and any(pd.notna(r) and str(r).strip() for r in result_comments):
        p_res = doc.add_paragraph()
        p_res.paragraph_format.line_spacing = 1
        for res in result_comments:
            if not pd.notna(res): continue
            texto = str(res).strip()
            if texto.lower() == "nan" or texto == "": continue
            
            run_item = p_res.add_run(texto + "\n")
            run_item.font.name = cfg['font_name']
            run_item.font.size = Pt(cfg['body'])
            
            if "not as expected" in texto.lower():
                run_item.bold = True
    else:
        doc.add_paragraph("")
    
    table_info = doc.add_table(rows=2, cols=3)
    format_cell(table_info.cell(0,0), "Witness", "")
    format_cell(table_info.cell(0,1), "Witness", "")
    format_cell(table_info.cell(0,2), "Date", "")
    
    def format_val(c, val):
        p = c.paragraphs[0]
        r = p.add_run(str(val) if pd.notna(val) else "")
        r.font.name = cfg['font_name']
        r.font.size = Pt(cfg['body'])
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        
    format_val(table_info.cell(1,0), test_info['Witness 1'])
    format_val(table_info.cell(1,1), test_info['Witness 2'])
    format_val(table_info.cell(1,2), test_info['Date:'])
    
    for row in table_info.rows:
        for cell in row.cells:
            set_cell_border(cell, **default_border_settings)

def generate_test_report_docx(excel_file, cfg):
    sheet_val = cfg['sheet_target']
    try:
        sheet_target = int(sheet_val)
    except ValueError:
        sheet_target = sheet_val
        
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_target, header=0)
    except ValueError:
        st.error(f"❌ Erro: Não foi possível encontrar a aba '{sheet_target}' na planilha. Verifique o nome ou índice no menu lateral.")
        return None
    except Exception as e:
        st.error(f"❌ Erro ao ler Excel: {e}")
        return None

    grouped_tests = {}
    for _, row in df.iterrows():
        if pd.isna(row.get('test number')): continue
        test_number = row['test number']
        if test_number not in grouped_tests:
            grouped_tests[test_number] = {
                'Test': row.get('Test', ''),
                'Method': row.get('Method', ''),
                'Steps': [],
                'Expected Results': [],
                'Result + Comment': [],
                'Max. Position Deviation (meters)': row.get('Max. Position Deviation (meters)', ''),
                'Max. Heading Deviation (degrees)': row.get('Max. Heading Deviation (degrees)', ''),
                'Witness 1': row.get('Witness 1', ''),
                'Witness 2': row.get('Witness 2', ''),
                'Date:': row.get('Date:', ''),
                'Section': row.get('Section', 'General')
            }
        grouped_tests[test_number]['Steps'].append(row.get('Step', ''))
        grouped_tests[test_number]['Expected Results'].append(row.get('Expected Result', ''))
        grouped_tests[test_number]['Result + Comment'].append(row.get('Result + Comment', ''))

    doc = Document()
    
    # ----------------------------------------------------
    # INSERÇÃO DA LOGO NO CABEÇALHO DO WORD
    # ----------------------------------------------------
    if os.path.exists(LOGO_DOC_PATH):
        try:
            # Acessa o cabeçalho da primeira seção (padrão)
            section = doc.sections[0]
            header = section.header
            p_header = header.paragraphs[0]
            r_header = p_header.add_run()
            # Ajusta tamanho da logo (1.5 polegadas é um bom padrão)
            r_header.add_picture(LOGO_DOC_PATH, width=Inches(1.5))
        except Exception as e:
            st.warning(f"Não foi possível inserir a logo no Word: {e}")

    # Configuração de Estilo Global
    style = doc.styles['Normal']
    style.font.name = cfg['font_name']
    style.font.size = Pt(cfg['body'])
    
    # Configuração de Margens
    for section in doc.sections:
        section.top_margin = Inches(cfg['m_top'])
        section.bottom_margin = Inches(cfg['m_bottom'])
        section.left_margin = Inches(cfg['m_left'])
        section.right_margin = Inches(cfg['m_right'])
        
    current_chapter = None
    for test_info in grouped_tests.values():
        if current_chapter != test_info['Section']:
            create_chapter_cover(doc, str(test_info['Section']), cfg)
            current_chapter = test_info['Section']
        create_test_page(doc, test_info, cfg)
        
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 4. FUNÇÕES DO MESCLADOR (PDF)
# ==========================================

def read_pdf_overlay_params(excel_file, sheet_target):
    try:
        try:
            st_target = int(sheet_target)
        except:
            st_target = sheet_target
        df = pd.read_excel(excel_file, sheet_name=st_target)
        nome_barco = df["Vessel"].iloc[0] if "Vessel" in df.columns else "Vessel Name"
        tipo_teste = df["Type"].iloc[0] if "Type" in df.columns else "Trials"
        mes_ano = df["Year"].iloc[0] if "Year" in df.columns else "202X"
        rodape_dir = df["Abreviation"].iloc[0] if "Abreviation" in df.columns else "DOC"
        rodape_esq = "Bram DP Assurance"
        rodape_center = ""
        return nome_barco, tipo_teste, mes_ano, rodape_esq, rodape_center, rodape_dir
    except Exception as e:
        st.error(f"Erro ao ler parâmetros do Excel para o PDF: {e}")
        return None

def create_overlay(page_width, page_height, page_number, params, font_name="Helvetica"):
    (nome_barco, tipo_teste, mes_ano, rodape_esq, rodape_center, rodape_dir) = params
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    
    margem = 40
    altura_cabecalho = page_height - 30
    
    font_pdf = "Helvetica" 
    if font_name in ["Times-Roman", "Courier", "Helvetica"]:
        font_pdf = font_name
        
    # ----------------------------------------------------
    # INSERÇÃO DA LOGO NO OVERLAY DO PDF
    # ----------------------------------------------------
    if os.path.exists(LOGO_DOC_PATH):
        try:
            # Desenha a logo no canto superior esquerdo
            # Ajuste as coordenadas (x, y, width, height) conforme necessário
            # Ex: x=40 (margem), y=altura_cabecalho - 15 (para ficar alinhado)
            logo_width = 80
            logo_height = 30
            c.drawImage(LOGO_DOC_PATH, x=margem, y=altura_cabecalho - 5, width=logo_width, height=logo_height, preserveAspectRatio=True, mask='auto')
            
            # Ajusta o texto do nome do barco para não ficar em cima da logo
            c.setFont(font_pdf, 12)
            # Desloca o nome do barco um pouco para a direita
            c.drawString(margem + logo_width + 10, altura_cabecalho, str(nome_barco))
            
        except Exception as e:
            # Fallback se a imagem der erro, desenha só o texto normal
            c.setFont(font_pdf, 12)
            c.drawString(margem, altura_cabecalho, str(nome_barco))
    else:
        # Sem logo, desenha normal
        c.setFont(font_pdf, 12)
        c.drawString(margem, altura_cabecalho, str(nome_barco))

    c.drawCentredString(page_width / 2, altura_cabecalho, str(tipo_teste))
    c.drawRightString(page_width - margem, altura_cabecalho, str(mes_ano))
    
    c.setFont(font_pdf, 10)
    altura_rodape = 20
    c.drawString(margem, altura_rodape, str(rodape_esq))
    c.drawCentredString(page_width / 2, altura_rodape, f"{rodape_center} {page_number}")
    c.drawRightString(page_width - margem, altura_rodape, str(rodape_dir))
    
    c.save()
    packet.seek(0)
    return PdfReader(packet)

def processar_pdf_final(doc1_bytes, doc2_bytes, excel_bytes, cfg):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as t1, \
         tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as t2:
        t1.write(doc1_bytes.read())
        t2.write(doc2_bytes.read())
        path1, path2 = t1.name, t2.name

    params = read_pdf_overlay_params(excel_bytes, cfg['sheet_target'])
    if not params: return None
    
    reader2 = PdfReader(path2)
    writer2 = PdfWriter()
    for page in reader2.pages:
        txt = page.extract_text()
        if txt and txt.strip():
            writer2.add_page(page)
    
    path2_clean = path2 + "_clean.pdf"
    with open(path2_clean, "wb") as f:
        writer2.write(f)
        
    merger = PdfMerger()
    merger.append(path1)
    merger.append(path2_clean)
    
    path_merged = path1 + "_merged.pdf"
    merger.write(path_merged)
    merger.close()
    
    reader_merged = PdfReader(path_merged)
    writer_final = PdfWriter()
    
    for i, page in enumerate(reader_merged.pages, start=1):
        pw = float(page.mediabox.width)
        ph = float(page.mediabox.height)
        overlay = create_overlay(pw, ph, i, params)
        page.merge_page(overlay.pages[0])
        writer_final.add_page(page)
        
    final_buffer = io.BytesIO()
    writer_final.write(final_buffer)
    final_buffer.seek(0)
    
    for p in [path1, path2, path2_clean, path_merged]:
        if os.path.exists(p): os.remove(p)
        
    return final_buffer

# ==========================================
# 5. INTERFACE PRINCIPAL
# ==========================================

# Define a logo do SITE
LOGO_SITE_PATH = "bram_logo.png"

col_logo1, col_logo2, col_logo3 = st.columns([1, 2, 1])

# Verifica se a logo do site existe para exibição
if os.path.exists(LOGO_SITE_PATH):
    with col_logo2:
        st.image(LOGO_SITE_PATH, use_container_width=True)
else:
    # Se não achar a bram_logo, tenta a logo.png como fallback para não ficar vazio
    if os.path.exists(LOGO_DOC_PATH):
         with col_logo2:
            st.image(LOGO_DOC_PATH, use_container_width=True)

st.markdown("<h1 style='text-align: center;'>Gerador de Relatórios DP</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: gray;'>Bram Offshore | Edison Chouest Offshore</p>", unsafe_allow_html=True)
st.markdown("---")

# --- ABAS ---
tab1, tab2 = st.tabs(["📄 1. Gerar Relatório (DOCX)", "📑 2. Mesclar e Finalizar (PDF)"])

with tab1:
    st.info("Passo 1: Faça upload da planilha preenchida para gerar o relatório em Word.")
    uploaded_excel = st.file_uploader("Upload Planilha Excel (.xlsx)", type=["xlsx"], key="u_excel_docx")
    
    if uploaded_excel:
        if st.button("Gerar Relatório DOCX", type="primary"):
            with st.spinner("Processando dados e gerando documento..."):
                docx_buffer = generate_test_report_docx(uploaded_excel, CONFIG)
                if docx_buffer:
                    st.success("Relatório gerado com sucesso!")
                    st.download_button(
                        label="⬇️ Baixar test_report.docx",
                        data=docx_buffer,
                        file_name="test_report.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

with tab2:
    st.info("Passo 2: Junte a parte inicial (Capa/Intro) com o relatório gerado acima.")
    col_up1, col_up2 = st.columns(2)
    pdf1 = col_up1.file_uploader("Parte 1 (Capa/Intro .pdf)", type=["pdf"])
    pdf2 = col_up2.file_uploader("Parte 2 (Relatório Gerado .pdf)", type=["pdf"], help="Converta o DOCX gerado na aba anterior para PDF antes de subir aqui.")
    excel_params = st.file_uploader("Planilha Excel (Para pegar Nome do Barco/Ano)", type=["xlsx"], key="u_excel_pdf")
    
    if pdf1 and pdf2 and excel_params:
        if st.button("Mesclar e Adicionar Cabeçalhos", type="primary"):
            with st.spinner("Mesclando arquivos e aplicando layout..."):
                final_pdf = processar_pdf_final(pdf1, pdf2, excel_params, CONFIG)
                if final_pdf:
                    st.success("PDF Final pronto!")
                    st.download_button(
                        label="⬇️ Baixar Relatório Final.pdf",
                        data=final_pdf,
                        file_name="Relatorio_Final_DP.pdf",
                        mime="application/pdf"
                    )
