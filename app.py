import os
import io
import sys
import tempfile
import streamlit as st
import pandas as pd
from PIL import Image

# -------------------------------
# IMPORTS PARA DOCX
# -------------------------------
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx2pdf import convert

# Função auxiliar para configurar as bordas de uma célula
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

# Configuração padrão de bordas
default_border_settings = {
    "top": {"sz": 4, "val": "single", "color": "000000", "space": "0"},
    "bottom": {"sz": 4, "val": "single", "color": "000000", "space": "0"},
    "left": {"sz": 0, "val": "nil", "color": "auto", "space": "0"},
    "right": {"sz": 0, "val": "nil", "color": "auto", "space": "0"}
}

def create_chapter_cover(doc, chapter_title):
    # Se já houver conteúdo, força um salto de página
    if len(doc.paragraphs) > 0:
        doc.add_page_break()
    # Adiciona 20 parágrafos vazios para o espaçamento superior
    for _ in range(20):
        p = doc.add_paragraph("")
        p.paragraph_format.line_spacing = 1
    # Cria o parágrafo do título centralizado
    para = doc.add_paragraph()
    para.paragraph_format.line_spacing = 1
    run = para.add_run(chapter_title)
    run.font.size = Pt(20)
    run.font.underline = True
    run.bold = True
    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    # Adiciona 7 parágrafos vazios para espaçamento inferior
    for _ in range(7):
        p = doc.add_paragraph("")
        p.paragraph_format.line_spacing = 1
    doc.add_page_break()

def create_bordered_section(doc, label, content, no_bottom_border=False, extra_space_top=Pt(3), extra_space_after_content=Pt(3)):
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    cell.text = ''
    # Parágrafo do título
    p_title = cell.add_paragraph()
    p_title.paragraph_format.space_before = extra_space_top
    p_title.paragraph_format.space_after = Pt(0)
    run_title = p_title.add_run(f"{label}:")
    run_title.bold = True
    run_title.font.size = Pt(12)
    # Parágrafo do conteúdo com espaçamento configurado
    p_content = cell.add_paragraph()
    p_content.paragraph_format.space_before = Pt(0)
    p_content.paragraph_format.space_after = extra_space_after_content
    run_content = p_content.add_run(content)
    run_content.font.size = Pt(11)
    if label.lower() == "method":
        p_content.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        cell.add_paragraph("")
    cell_border_settings = default_border_settings.copy()
    if no_bottom_border:
        cell_border_settings.pop("bottom", None)
    set_cell_border(cell, **cell_border_settings)

def create_test_page(doc, test_info):
    # Inicia nova seção para cada página de teste
    doc.add_section(WD_SECTION_START.NEW_PAGE)
    # Título do teste
    para_title = doc.add_paragraph()
    para_title.paragraph_format.line_spacing = 1
    run_title = para_title.add_run(f"{test_info['Test']}")
    run_title.bold = True
    run_title.font.size = Pt(14)
    para_title.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    doc.add_paragraph("")
    # Seção "Method"
    create_bordered_section(doc, "Method", test_info['Method'], extra_space_top=Pt(1), extra_space_after_content=Pt(3))
    doc.add_paragraph("")
    # Seção "Steps"
    para_steps_title = doc.add_paragraph()
    para_steps_title.paragraph_format.line_spacing = 1
    run_steps_title = para_steps_title.add_run("Steps:")
    run_steps_title.bold = True
    run_steps_title.font.size = Pt(12)
    for step in test_info['Steps']:
        p_step = doc.add_paragraph(f"{step}")
        p_step.paragraph_format.line_spacing = 1
    doc.add_paragraph("")
    # Seção "Expected Results"
    create_bordered_section(doc, "Expected Results", "\n".join(test_info['Expected Results']),
                              no_bottom_border=True, extra_space_top=Pt(1), extra_space_after_content=Pt(0))
    doc.add_paragraph("")
    # Tabela com informações adicionais
    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    cell_1 = table.cell(0, 0)
    p1 = cell_1.paragraphs[0]
    p1.paragraph_format.line_spacing = 1
    run1 = p1.add_run("Max Deviation:")
    run1.bold = True
    p1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    cell_2 = table.cell(0, 1)
    p2 = cell_2.paragraphs[0]
    p2.paragraph_format.line_spacing = 1
    run2_label = p2.add_run("Position:")
    run2_label.bold = True
    run2_content = p2.add_run(f" < {test_info['Max. Position Deviation (meters)']} meters")
    p2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    cell_3 = table.cell(0, 2)
    p3 = cell_3.paragraphs[0]
    p3.paragraph_format.line_spacing = 1
    run3_label = p3.add_run("Heading:")
    run3_label.bold = True
    run3_content = p3.add_run(f" < {test_info['Max. Heading Deviation (degrees)']} degrees")
    p3.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell, **default_border_settings)
    doc.add_paragraph("")
    # Seção "Results"
    para_res_title = doc.add_paragraph()
    para_res_title.paragraph_format.line_spacing = 1
    run_res_title = para_res_title.add_run("Results:")
    run_res_title.bold = True
    run_res_title.font.size = Pt(12)
    result_comments = test_info.get('Result + Comment')
    if result_comments and any(pd.notna(r) and str(r).strip() for r in result_comments):
        results_paragraph = doc.add_paragraph()
        results_paragraph.paragraph_format.line_spacing = 1
        for res in result_comments:
            if not pd.notna(res):
                continue
            texto = str(res).strip()
            if texto.lower() == "nan" or texto == "":
                continue
            if "not as expected" in texto.lower():
                run_item = results_paragraph.add_run(texto + "\n")
                run_item.bold = True
            else:
                results_paragraph.add_run(texto + "\n")
    else:
        doc.add_paragraph("")
    # Tabela final com Witness e Date
    table_info = doc.add_table(rows=2, cols=3)
    table_info.style = None
    hdr_cells = table_info.rows[0].cells
    p_hdr0 = hdr_cells[0].paragraphs[0]
    p_hdr0.paragraph_format.line_spacing = 1
    run_hdr1 = p_hdr0.add_run("Witness")
    run_hdr1.bold = True
    p_hdr0.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p_hdr1 = hdr_cells[1].paragraphs[0]
    p_hdr1.paragraph_format.line_spacing = 1
    run_hdr2 = p_hdr1.add_run("Witness")
    run_hdr2.bold = True
    p_hdr1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p_hdr2 = hdr_cells[2].paragraphs[0]
    p_hdr2.paragraph_format.line_spacing = 1
    run_hdr3 = p_hdr2.add_run("Date")
    run_hdr3.bold = True
    p_hdr2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    val_cells = table_info.rows[1].cells
    p_val0 = val_cells[0].paragraphs[0]
    p_val0.paragraph_format.line_spacing = 1
    p_val0.add_run(str(test_info['Witness 1']))
    p_val0.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p_val1 = val_cells[1].paragraphs[0]
    p_val1.paragraph_format.line_spacing = 1
    p_val1.add_run(str(test_info['Witness 2']))
    p_val1.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p_val2 = val_cells[2].paragraphs[0]
    p_val2.paragraph_format.line_spacing = 1
    p_val2.add_run(str(test_info['Date:']))
    p_val2.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    for row in table_info.rows:
        for cell in row.cells:
            set_cell_border(cell, **default_border_settings)

def generate_test_report_docx(excel_path):
    df = pd.read_excel(excel_path, sheet_name=1, header=0)
    grouped_tests = {}
    for _, row in df.iterrows():
        chapter_title = row['Section']
        test_number = row['test number']
        if test_number not in grouped_tests:
            grouped_tests[test_number] = {
                'Test': row['Test'],
                'Method': row['Method'],
                'Steps': [],
                'Expected Results': [],
                'Result + Comment': [],
                'Max. Position Deviation (meters)': row['Max. Position Deviation (meters)'],
                'Max. Heading Deviation (degrees)': row['Max. Heading Deviation (degrees)'],
                'Witness 1': row['Witness 1'],
                'Witness 2': row['Witness 2'],
                'Date:': row['Date:'],
                'Section': chapter_title
            }
        grouped_tests[test_number]['Steps'].append(row['Step'])
        grouped_tests[test_number]['Expected Results'].append(row['Expected Result'])
        grouped_tests[test_number]['Result + Comment'].append(row['Result + Comment'])
    doc = Document()
    # Configuração do estilo "Normal" – igual à versão usada no Colab
    styles = doc.styles
    normal_style = styles['Normal']
    normal_style.font.name = 'Raleway'
    normal_paragraph_format = normal_style.paragraph_format
    normal_paragraph_format.space_before = Pt(0)
    normal_paragraph_format.space_after = Pt(0)
    normal_paragraph_format.line_spacing = 1
    # Ajusta margens das seções
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(0.2)
        section.right_margin = Inches(0.2)
    current_chapter = None
    for test_info in grouped_tests.values():
        if current_chapter != test_info['Section']:
            create_chapter_cover(doc, test_info['Section'])
            current_chapter = test_info['Section']
        create_test_page(doc, test_info)
    output_docx = os.path.join(tempfile.gettempdir(), "test_report.docx")
    doc.save(output_docx)
    return output_docx

def convert_docx_to_pdf(docx_path):
    pdf_path = docx_path.replace(".docx", ".pdf")
    convert(docx_path, pdf_path)
    return pdf_path

# -------------------------------
# IMPORTS E FUNÇÕES PARA PDF
# -------------------------------
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(TTFont('Raleway', 'Raleway-Regular.ttf'))

def read_pdf_overlay_params(excel_path):
    df = pd.read_excel(excel_path, sheet_name=1)
    nome_barco = df.loc[0, "Vessel"]
    tipo_teste = df.loc[0, "Type"]
    mes_ano = df.loc[0, "Year"]
    rodape_dir = df.loc[0, "Abreviation"]
    rodape_esq = "Bram DP Assurance"
    rodape_center = ""
    return nome_barco, tipo_teste, mes_ano, rodape_esq, rodape_center, rodape_dir

def create_overlay(page_width, page_height, page_number, overlay_params):
    (nome_barco, tipo_teste, mes_ano, rodape_esq, rodape_center, rodape_dir) = overlay_params
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))
    margem = 40
    altura_cabecalho = page_height - 30
    c.setFont("Raleway", 12)
    c.drawString(margem, altura_cabecalho, nome_barco)
    c.drawCentredString(page_width / 2, altura_cabecalho, tipo_teste)
    c.drawRightString(page_width - margem, altura_cabecalho, mes_ano)
    c.setFont("Raleway", 10)
    altura_rodape = 20
    c.drawString(margem, altura_rodape, rodape_esq)
    c.drawCentredString(page_width / 2, altura_rodape, f"{rodape_center} {page_number}")
    c.drawRightString(page_width - margem, altura_rodape, rodape_dir)
    c.save()
    packet.seek(0)
    return PdfReader(packet)

def add_header_footer(input_pdf_path, output_pdf_path, overlay_params):
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    for i, page in enumerate(reader.pages, start=1):
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        overlay_pdf = create_overlay(page_width, page_height, i, overlay_params)
        overlay_page = overlay_pdf.pages[0]
        page.merge_page(overlay_page)
        writer.add_page(page)
    with open(output_pdf_path, "wb") as f_out:
        writer.write(f_out)
    return output_pdf_path

def merge_pdfs_func(pdf_list, output_pdf_path):
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(pdf)
    merger.write(output_pdf_path)
    merger.close()
    return output_pdf_path

def remove_blank_pages(input_pdf_path, output_pdf_path):
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()
    for i, page in enumerate(reader.pages, start=1):
        text = page.extract_text()
        if text and text.strip():
            writer.add_page(page)
    with open(output_pdf_path, "wb") as f_out:
        writer.write(f_out)
    return output_pdf_path

def run_pdf_merge(doc1_path, doc2_path, excel_path):
    pdf_relatorio_sem_blank = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
    remove_blank_pages(doc2_path, pdf_relatorio_sem_blank)
    merged_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
    merge_pdfs_func([doc1_path, pdf_relatorio_sem_blank], merged_pdf)
    overlay_params = read_pdf_overlay_params(excel_path)
    final_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf").name
    add_header_footer(merged_pdf, final_pdf, overlay_params)
    return final_pdf

# -------------------------------
# INTERFACE COM STREAMLIT
# -------------------------------

st.title("Sistema de Relatórios e Mesclagem de PDFs")

# -------------------------------
# EXEMPLO: ADICIONANDO UMA IMAGEM NA INTERFACE
# -------------------------------
# Para exibir uma imagem, você precisa saber o caminho (path) onde ela está salva.
# Em um Mac, o caminho completo geralmente segue o padrão:
#   /Users/seu_usuario/caminho/para/imagem.png
#
# Dicas para obter o caminho da imagem no Mac:
# 1. Abra o Finder e localize a imagem.
# 2. Clique com o botão direito na imagem e escolha "Obter Informações".
# 3. Na janela de informações, procure por "Onde:" que mostrará o caminho para a pasta.
# 4. Combine esse caminho com o nome do arquivo para formar o caminho completo.
#
# Exemplo: Se seu usuário for "joao" e a imagem estiver na pasta Downloads com o nome "minha_imagem.png",
# o caminho seria: "/Users/joao/Downloads/minha_imagem.png"
#
# Altere o valor de 'caminho_imagem' para o caminho correto da sua imagem.
caminho_imagem = "images/Logo tradicional.png"

try:
    imagem = Image.open(caminho_imagem)
    # Utilizando colunas para centralizar a imagem (opcional)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.image(imagem, caption="Imagem Exemplo", use_column_width=True)
except Exception as e:
    st.error(f"Erro ao carregar a imagem. Verifique o caminho. Detalhes: {e}")

# -------------------------------
# Abas do aplicativo
# -------------------------------
tab1, tab2 = st.tabs(["Gerar Relatório", "Mesclar PDFs"])

with tab1:
    st.header("1. Gerar Test Report")
    excel_file = st.file_uploader("Upload da planilha Excel", type=["xlsx"], key="excel_gen")
    if excel_file:
        if st.button("Gerar DOCX"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(excel_file.read())
                tmp_excel = tmp.name
            docx_path = generate_test_report_docx(tmp_excel)
            st.success("DOCX gerado!")
            with open(docx_path, "rb") as f:
                st.download_button("Baixar DOCX", f, file_name="test_report.docx")
       

with tab2:
    st.header("2. Mesclar PDFs")
    doc1 = st.file_uploader("Upload do PDF da primeira metade (doc1.pdf)", type=["pdf"], key="pdf1")
    doc2 = st.file_uploader("Upload do PDF do Test Report (doc2.pdf)", type=["pdf"], key="pdf2")
    excel_for_pdf = st.file_uploader("Upload da planilha Excel (para parâmetros)", type=["xlsx"], key="excel_pdf")
    if st.button("Mesclar PDFs"):
        if doc1 and doc2 and excel_for_pdf:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp1:
                tmp1.write(doc1.read())
                doc1_path = tmp1.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp2:
                tmp2.write(doc2.read())
                doc2_path = tmp2.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_excel:
                tmp_excel.write(excel_for_pdf.read())
                excel_path = tmp_excel.name
            final_pdf = run_pdf_merge(doc1_path, doc2_path, excel_path)
            st.success("PDF final mesclado!")
            with open(final_pdf, "rb") as f:
                st.download_button("Baixar PDF Final", f, file_name="final_report.pdf")
