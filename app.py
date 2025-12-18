import streamlit as st
import pandas as pd
import re
import os
import io
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==========================================
# CONFIGURAÇÃO DA PÁGINA STREAMLIT
# ==========================================
st.set_page_config(page_title="Gerador de Relatórios DP", layout="wide", page_icon="⚓")

# Estilo CSS para manter o padrão visual azul
st.markdown("""
    <style>
    :root { --primary-color: #1F4E79; }
    div.stButton > button:first-child {
        background-color: #1F4E79;
        color: white;
        border-radius: 5px;
        border: none;
        font-weight: bold;
    }
    div.stButton > button:hover { background-color: #163a5c; color: white; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# CONSTANTES E CONFIGURAÇÕES DO DOCX
# ==========================================
COLOR_PRIMARY = "1F4E79"
COLOR_BORDER  = "BFBFBF"
COLOR_BG_UNIFIED = "F2F2F2"
COLOR_TEXT_MAIN = RGBColor(0x26, 0x26, 0x26)
COLOR_TEXT_LABEL = RGBColor(0x1F, 0x4E, 0x79)
COLOR_TEXT_PLACEHOLDER = RGBColor(89, 89, 89)

# Define o caminho da logo no repositório
LOGO_PATH = 'logo.png' 

# Configurações de Borda
refined_border = {"sz": 8, "val": "single", "color": COLOR_BORDER}
box_border_settings = {
    "top": refined_border, "bottom": refined_border, "left": refined_border, "right": refined_border
}

no_border = {
    "top": {"sz": 0, "val": "nil", "color": "auto"},
    "bottom": {"sz": 0, "val": "nil", "color": "auto"},
    "left": {"sz": 0, "val": "nil", "color": "auto"},
    "right": {"sz": 0, "val": "nil", "color": "auto"},
    "insideV": {"sz": 0, "val": "nil", "color": "auto"}
}

# ==========================================
# FUNÇÕES UTILITÁRIAS (FORMATAÇÃO WORD)
# ==========================================

def set_cell_border_and_shading(cell, border_settings=None, shading_color=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    for element in tcPr.xpath('./w:tcBorders'):
        tcPr.remove(element)

    if border_settings:
        tcBorders = OxmlElement('w:tcBorders')
        for edge, data in border_settings.items():
            element = OxmlElement(f"w:{edge}")
            for key, value in data.items():
                element.set(qn(f"w:{key}"), str(value))
            tcBorders.append(element)
        tcPr.append(tcBorders)

    for element in tcPr.xpath('./w:shd'):
        tcPr.remove(element)

    if shading_color:
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:fill'), shading_color)
        tcPr.append(shd)

def set_cell_margins(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, value in kwargs.items():
        mar = OxmlElement(f'w:{edge}')
        mar.set(qn('w:w'), str(value))
        mar.set(qn('w:type'), 'dxa')
        tcMar.append(mar)
    tcPr.append(tcMar)

def set_table_indent(table, indent_val=0):
    tblPr = table._tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        table._tbl.insert(0, tblPr)

    # Layout Fixo
    layout = OxmlElement('w:tblLayout')
    layout.set(qn('w:type'), 'fixed')
    for el in tblPr.xpath("w:tblLayout"): tblPr.remove(el)
    tblPr.append(layout)

    # Remover espaçamento
    for element in tblPr.xpath('./w:tblCellSpacing'): tblPr.remove(element)
    tblCellSpacing = OxmlElement('w:tblCellSpacing')
    tblCellSpacing.set(qn('w:w'), "0")
    tblCellSpacing.set(qn('w:type'), "dxa")
    tblPr.append(tblCellSpacing)

    # Indentação Controlada
    for el in tblPr.xpath("w:tblInd"): tblPr.remove(el)
    tblInd = OxmlElement('w:tblInd')
    tblInd.set(qn('w:w'), str(indent_val))
    tblInd.set(qn('w:type'), 'dxa')
    tblPr.append(tblInd)

def create_header(doc, image_path):
    section = doc.sections[0]
    section.header_distance = Inches(0.2)
    header = section.header

    for paragraph in header.paragraphs:
        p_element = paragraph._element
        p_element.getparent().remove(p_element)

    p_logo = header.add_paragraph()
    p_logo.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    p_logo.paragraph_format.space_before = Pt(0)
    p_logo.paragraph_format.space_after = Pt(0)
    p_logo.paragraph_format.left_indent = Inches(-0.09)

    if os.path.exists(image_path):
        run_logo = p_logo.add_run()
        run_logo.add_picture(image_path, width=Inches(7.5))
    else:
        # Fallback caso a logo não exista, para não quebrar o código
        run_logo = p_logo.add_run("[LOGO NÃO ENCONTRADO - Verifique se logo.png está no repositório]")
        run_logo.font.color.rgb = RGBColor(255, 0, 0)
        run_logo.font.size = Pt(8)

def create_details_section(doc, test_info):
    details_table = doc.add_table(rows=0, cols=1)
    details_table.width = Inches(7.5)
    details_table.allow_autofit = False

    set_table_indent(details_table, indent_val=0)

    details_table.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    def _add_box_row(label, content, shading_color, is_placeholder=False):
        row = details_table.add_row()
        cell = row.cells[0]
        cell.width = Inches(7.5)
        set_cell_margins(cell, top=80, bottom=80, left=120, right=100)
        set_cell_border_and_shading(cell, border_settings=box_border_settings, shading_color=shading_color)
        if label:
            p_label = cell.paragraphs[0]
            if not p_label.text: p_label.clear()
            p_label.paragraph_format.space_before = Pt(0); p_label.paragraph_format.space_after = Pt(3)
            run_label = p_label.add_run(f"{label}:"); run_label.bold = True
            run_label.font.size = Pt(9); run_label.font.color.rgb = COLOR_TEXT_LABEL
        p_content = cell.add_paragraph() if label else cell.paragraphs[0]
        p_content.paragraph_format.space_before = Pt(0); p_content.paragraph_format.space_after = Pt(2)
        p_content.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        if is_placeholder:
            run_content = p_content.add_run(content)
            run_content.font.italic = True
            run_content.font.color.rgb = COLOR_TEXT_PLACEHOLDER
            run_content.font.size = Pt(9)
        else:
            lines = str(content).split('\n')
            for i, line in enumerate(lines):
                line = line.strip()
                if not line: continue
                match = re.match(r'^(\d+\.)\s*(.*)', line)
                if match:
                    num_part = match.group(1); text_part = match.group(2)
                    r_num = p_content.add_run(num_part + " "); r_num.font.size = Pt(9); r_num.font.color.rgb = COLOR_TEXT_MAIN; r_num.bold = True
                    r_txt = p_content.add_run(text_part); r_txt.font.size = Pt(9); r_txt.font.color.rgb = COLOR_TEXT_MAIN; r_txt.bold = False
                else:
                    r_line = p_content.add_run(line); r_line.font.size = Pt(9); r_line.font.color.rgb = COLOR_TEXT_MAIN
                if i < len(lines) - 1: p_content.add_run("\n")
        return p_content

    if test_info.get('Objective'):
        _add_box_row("Objective", str(test_info['Objective']), COLOR_BG_UNIFIED)

    _add_box_row("Method", test_info['Method'], COLOR_BG_UNIFIED)

    steps_content = "\n".join(str(step) for step in test_info['Steps'])
    _add_box_row("Steps", steps_content, COLOR_BG_UNIFIED)

    _add_box_row("Expected Results", "\n".join(map(str, test_info['Expected Results'])), COLOR_BG_UNIFIED)

    results_content_str = "\n".join(str(res).strip() for res in test_info.get('Result + Comment', []) if pd.notna(res) and str(res).strip() and str(res).strip().lower() != "nan")
    if not results_content_str:
        results_content_str = "No results or comments provided."
        is_ph = True
    else:
        is_ph = False
    _add_box_row("Results", results_content_str, COLOR_BG_UNIFIED, is_placeholder=is_ph)

    comments_list = test_info.get('Step Comments', [])
    row = details_table.add_row()
    cell = row.cells[0]; cell.width = Inches(7.5)
    set_cell_margins(cell, top=80, bottom=80, left=120, right=100)
    set_cell_border_and_shading(cell, border_settings=box_border_settings, shading_color=COLOR_BG_UNIFIED)

    p_label = cell.paragraphs[0]
    if not p_label.text: p_label.clear()
    p_label.paragraph_format.space_before = Pt(0); p_label.paragraph_format.space_after = Pt(3)
    run_label = p_label.add_run("Comments:"); run_label.bold = True; run_label.font.size = Pt(9); run_label.font.color.rgb = COLOR_TEXT_LABEL
    p_content = cell.add_paragraph()
    p_content.paragraph_format.space_before = Pt(0); p_content.paragraph_format.space_after = Pt(2)
    p_content.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    if not comments_list:
        run_ph = p_content.add_run("No additional comments")
        run_ph.font.italic = True; run_ph.font.size = Pt(9)
        run_ph.font.color.rgb = COLOR_TEXT_PLACEHOLDER
    else:
        for i, item in enumerate(comments_list):
            run_step = p_content.add_run(f"Step {item['step']}: "); run_step.bold = True; run_step.font.size = Pt(9); run_step.font.color.rgb = COLOR_TEXT_MAIN
            run_text = p_content.add_run(f"{item['text']}"); run_text.bold = False; run_text.font.size = Pt(9); run_text.font.color.rgb = COLOR_TEXT_MAIN
            if i < len(comments_list) - 1: p_content.add_run("\n")

def create_test_page(doc, test_info, is_first_test=False, section_title=None):
    if not is_first_test and len(doc.paragraphs) > 0:
        doc.add_page_break()

    if section_title:
        p_sec = doc.add_paragraph()
        p_sec.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        p_sec.paragraph_format.space_before = Pt(6); p_sec.paragraph_format.space_after = Pt(14)
        run_sec = p_sec.add_run(str(section_title))
        run_sec.font.name = 'Raleway'; run_sec.font.size = Pt(16); run_sec.font.color.rgb = RGBColor(0x0, 0x0, 0x0); run_sec.bold = True

    # --- TABELA AZUL (HEADER) ---
    table_blue = doc.add_table(rows=1, cols=2)
    table_blue.width = Inches(7.5)
    table_blue.allow_autofit = False

    # Indentação negativa para compensar a falta de borda
    set_table_indent(table_blue, indent_val=-10)

    table_blue.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    table_blue.columns[0].width = Inches(0.8)
    table_blue.columns[1].width = Inches(6.7)

    cell_lbl = table_blue.cell(0, 0)
    cell_val = table_blue.cell(0, 1)
    cell_lbl.width = Inches(0.8)
    cell_val.width = Inches(6.7)

    set_cell_margins(cell_lbl, top=60, bottom=60, left=100, right=100)
    set_cell_margins(cell_val, top=60, bottom=60, left=100, right=100)

    p_lbl = cell_lbl.paragraphs[0]
    run_lbl = p_lbl.add_run("TEST NO:")
    run_lbl.bold = True; run_lbl.font.size = Pt(10); run_lbl.font.color.rgb = RGBColor(255, 255, 255)
    p_lbl.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    p_val = cell_val.paragraphs[0]
    run_val = p_val.add_run(f"{test_info['Test']}")
    run_val.bold = True; run_val.font.size = Pt(10); run_val.font.color.rgb = RGBColor(255, 255, 255)
    p_val.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    set_cell_border_and_shading(cell_lbl, border_settings=no_border, shading_color=COLOR_PRIMARY)
    set_cell_border_and_shading(cell_val, border_settings=no_border, shading_color=COLOR_PRIMARY)

    p_gap_header = doc.add_paragraph()
    p_gap_header.paragraph_format.space_before = Pt(0)
    p_gap_header.paragraph_format.space_after = Pt(0)
    p_gap_header.paragraph_format.line_spacing = Pt(2)

    # --- TABELA FMEA ---
    table_info = doc.add_table(rows=1, cols=2)
    table_info.width = Inches(7.49)
    table_info.allow_autofit = False

    set_table_indent(table_info, indent_val=0)

    table_info.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    table_info.columns[0].width = Inches(3.75)
    table_info.columns[1].width = Inches(3.74)

    cell_fmea = table_info.cell(0, 0)
    cell_sub = table_info.cell(0, 1)

    cell_fmea.width = Inches(3.75)
    cell_sub.width = Inches(3.74)

    set_cell_margins(cell_fmea, top=60, bottom=60, left=100, right=100)
    set_cell_margins(cell_sub, top=60, bottom=60, left=100, right=100)

    p_fmea = cell_fmea.paragraphs[0]
    run_fmea = p_fmea.add_run("FMEA Reference: "); run_fmea.bold = True; run_fmea.font.size = Pt(9); run_fmea.font.color.rgb = COLOR_TEXT_MAIN
    p_fmea.add_run(str(test_info.get('FMEA Reference', '-'))).font.size = Pt(9); p_fmea.runs[-1].font.color.rgb = COLOR_TEXT_MAIN

    set_cell_border_and_shading(cell_fmea, border_settings=box_border_settings, shading_color=COLOR_BG_UNIFIED)

    p_sub = cell_sub.paragraphs[0]
    run_sub = p_sub.add_run("Sub-System: "); run_sub.bold = True; run_sub.font.size = Pt(9); run_sub.font.color.rgb = COLOR_TEXT_MAIN
    p_sub.add_run(str(test_info.get('Sub-System', '-'))).font.size = Pt(9); p_sub.runs[-1].font.color.rgb = COLOR_TEXT_MAIN

    set_cell_border_and_shading(cell_sub, border_settings=box_border_settings, shading_color=COLOR_BG_UNIFIED)

    create_details_section(doc, test_info)

    p_gap_footer = doc.add_paragraph()
    p_gap_footer.paragraph_format.space_before = Pt(0)
    p_gap_footer.paragraph_format.space_after = Pt(0)
    p_gap_footer.paragraph_format.line_spacing = Pt(2)

    # --- TABELA ASSINATURAS ---
    table_witness = doc.add_table(rows=1, cols=2)
    table_witness.width = Inches(7.5)
    table_witness.allow_autofit = False

    set_table_indent(table_witness, indent_val=-10)

    table_witness.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    table_witness.columns[0].width = Inches(5.0)
    table_witness.columns[1].width = Inches(2.5)

    cell_wit = table_witness.cell(0, 0)
    cell_dat = table_witness.cell(0, 1)

    cell_wit.width = Inches(5.0)
    cell_dat.width = Inches(2.5)

    set_cell_margins(cell_wit, top=60, bottom=60, left=100, right=100)
    set_cell_margins(cell_dat, top=60, bottom=60, left=100, right=100)

    p_wit = cell_wit.paragraphs[0]
    run_wit_lbl = p_wit.add_run("Witnessed by: "); run_wit_lbl.bold = True; run_wit_lbl.font.size = Pt(9); run_wit_lbl.font.color.rgb = RGBColor(255, 255, 255)
    run_wit_val = p_wit.add_run(str(test_info.get('Witness 1', '-'))); run_wit_val.bold = True; run_wit_val.font.size = Pt(9); run_wit_val.font.color.rgb = RGBColor(255, 255, 255)
    p_wit.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    p_dat = cell_dat.paragraphs[0]
    run_dat_lbl = p_dat.add_run("Date: "); run_dat_lbl.bold = True; run_dat_lbl.font.size = Pt(9); run_dat_lbl.font.color.rgb = RGBColor(255, 255, 255)

    raw_date = test_info.get('Date:')
    date_value = str(raw_date).strip() if pd.notna(raw_date) and str(raw_date).lower() != 'nan' else '-'

    run_dat_val = p_dat.add_run(date_value)
    run_dat_val.bold = True; run_dat_val.font.size = Pt(9); run_dat_val.font.color.rgb = RGBColor(255, 255, 255)
    p_dat.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    set_cell_border_and_shading(cell_wit, border_settings=no_border, shading_color=COLOR_PRIMARY)
    set_cell_border_and_shading(cell_dat, border_settings=no_border, shading_color=COLOR_PRIMARY)


def generate_professional_docx(uploaded_file):
    # 1. Ler o Excel
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0, header=0, dtype={'test number': str})
    except Exception as e:
        st.error(f"Erro ao ler o arquivo Excel: {e}")
        return None

    grouped_tests = {}
    if not df.empty:
        for _, row in df.iterrows():
            chapter_title = row.get('Section', 'Section Name')
            test_number = row.get('test number', '000')
            
            # Pula se não tiver número de teste
            if pd.isna(test_number): continue

            if test_number not in grouped_tests:
                grouped_tests[test_number] = {
                    'Test': row.get('Test', 'Test Title'), 
                    'Method': row.get('Method', ''), 
                    'Steps': [],
                    'Expected Results': [], 
                    'Result + Comment': [], 
                    'Step Comments': [],
                    'Witness 1': row.get('Witness 1', ''), 
                    'Date:': row.get('Date', ''), 
                    'Section': chapter_title,
                    'FMEA Reference': row.get('FMEA Reference', ''), 
                    'Sub-System': row.get('Sub-System', ''),
                    'Objective': row.get('Objective', '')
                }
            grouped_tests[test_number]['Steps'].append(row.get('Step', ''))
            
            # Comentários de Auditoria
            current_step_num = len(grouped_tests[test_number]['Steps'])
            auditor_comment = row.get('Auditor FMEA Comment')
            if pd.notna(auditor_comment) and str(auditor_comment).strip():
                grouped_tests[test_number]['Step Comments'].append({'step': current_step_num, 'text': str(auditor_comment).strip()})
            
            grouped_tests[test_number]['Expected Results'].append(row.get('Expected Result', ''))
            grouped_tests[test_number]['Result + Comment'].append(row.get('Result + Comment', ''))

    # 2. Criar Documento
    doc = Document()
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Raleway'
    normal_style.font.size = Pt(10)
    normal_style.paragraph_format.space_before = Pt(0)
    normal_style.paragraph_format.space_after = Pt(0)
    normal_style.paragraph_format.line_spacing = 1.15

    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)

    # Header com logo
    create_header(doc, LOGO_PATH)

    # Título Principal
    main_title_para = doc.add_paragraph()
    main_title_run = main_title_para.add_run("APPENDIX B - DP ANNUAL TRIALS TESTS")
    main_title_run.font.name = 'Raleway'
    main_title_run.font.size = Pt(20)
    main_title_run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    main_title_run.bold = True
    main_title_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    main_title_para.paragraph_format.space_after = Pt(24)

    first_iteration = True
    current_chapter = None
    
    if grouped_tests:
        for test_number, test_info in grouped_tests.items():
            test_info['test number'] = test_number
            if test_info['Section'] != current_chapter:
                current_chapter = test_info['Section']
                create_test_page(doc, test_info, is_first_test=first_iteration, section_title=current_chapter)
            else:
                create_test_page(doc, test_info, is_first_test=first_iteration, section_title=None)
            first_iteration = False

    # Salvar em buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# INTERFACE DO USUÁRIO
# ==========================================

# Logo na Interface Streamlit
if os.path.exists('bram_logo.png'):
     col1, col2, col3 = st.columns([1, 2, 1])
     with col2:
         st.image('bram_logo.png', use_container_width=True)

st.title("Gerador de Relatórios DP - Padrão Profissional")
st.markdown("Faça o upload da planilha Excel para gerar o relatório formatado (Apêndice B).")

uploaded_file = st.file_uploader("Upload Planilha (.xlsx)", type=["xlsx"])

if uploaded_file:
    if st.button("Gerar Relatório DOCX", type="primary"):
        with st.spinner("Processando formatação avançada..."):
            docx_buffer = generate_professional_docx(uploaded_file)
            
            if docx_buffer:
                st.success("Relatório gerado com sucesso!")
                st.download_button(
                    label="⬇️ Baixar test_report_professional.docx",
                    data=docx_buffer,
                    file_name="test_report_professional.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
