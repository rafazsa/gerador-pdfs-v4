import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import requests
import math
import os
import zipfile
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ===================== CONFIG VISUAL =====================
PRIMARY = colors.HexColor("#0F3460")
ACCENT  = colors.HexColor("#00A6A6")
TEXT    = colors.HexColor("#1F2937")
ZEBRA_1 = colors.whitesmoke
ZEBRA_2 = colors.HexColor("#ffffff")
BORDER  = colors.HexColor("#E5E7EB")

BASE_FONT = "Helvetica"
PAGE_SIZE = A4
PAGE_W, PAGE_H = PAGE_SIZE
MARGIN_TOP_MM = 24.0
MARGIN_BOTTOM_MM = 18.0
MARGIN_SIDE_MM = 14.0
REPORT_TITLE = "Tratamento de Sementes (Resumo por Registro)"
LOGO_URL = "https://report.geodata.com.br/customers/picture/16?time=1759945360"
LOGO_HEIGHT_MM = 15.0
IMG_MAX_W = 130 * mm
IMG_MAX_H = 70 * mm
SIGNATURE_MAX_H = 25 * mm  # altura máxima para imagem de assinatura

# ===================== ESTILOS =====================
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name="ReportTitle", fontName=BASE_FONT, fontSize=12.5, textColor=PRIMARY, alignment=1))
styles.add(ParagraphStyle(name="SectionTitle", fontName=BASE_FONT, fontSize=9.8, textColor=PRIMARY))
styles.add(ParagraphStyle(name="Q", fontName=BASE_FONT, fontSize=8.6, textColor=colors.HexColor("#111827")))
styles.add(ParagraphStyle(name="A", fontName=BASE_FONT, fontSize=8.6, textColor=colors.HexColor("#111827")))

# ===================== FUNÇÕES AUXILIARES =====================
def fetch_image(url, max_w=None, max_h=None, align_center=False):
    if not url or not str(url).strip():
        return None
    try:
        resp = requests.get(url, timeout=15)
        resp.raise_for_status()
        img = Image(BytesIO(resp.content))
        iw, ih = img.wrap(0, 0)
        if max_w or max_h:
            scale_w = (max_w / iw) if max_w else 1.0
            scale_h = (max_h / ih) if max_h else 1.0
            scale = min(scale_w, scale_h)
            img._restrictSize(iw * scale, ih * scale)
        if align_center:
            img.hAlign = "CENTER"
        return img
    except Exception:
        return None

def normalize_value(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def make_qa_table(pairs, pairs_per_row, available_width):
    q_frac, a_frac = 0.35, 0.65
    widths = []
    for _ in range(pairs_per_row):
        widths += [available_width * q_frac / pairs_per_row, available_width * a_frac / pairs_per_row]

    data_rows = []
    for q, a in pairs:
        q_par = Paragraph(q, styles["Q"])
        a_par = a if isinstance(a, Image) else Paragraph(a if a else "-", styles["A"])
        data_rows.append([q_par, a_par])

    table = Table(data_rows, colWidths=widths, hAlign="LEFT")
    table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.25, BORDER),
        ("ROWBACKGROUNDS", (0,0), (-1,-1), [ZEBRA_1, ZEBRA_2]),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
    ]))
    return table

def gerar_pdfs(file_bytes):
    df_raw = pd.read_excel(file_bytes, header=None)
    if df_raw.shape[0] < 3:
        raise ValueError("Planilha inesperada: preciso de pelo menos 3 linhas (título, cabeçalho e dados).")

    output_dir = "relatorios_individuais"
    os.makedirs(output_dir, exist_ok=True)
    generated_files = []

    for r_index in range(2, df_raw.shape[0]):
        row = df_raw.iloc[r_index].tolist()
        pairs = []
        user_name = normalize_value(row[2]) if len(row) > 2 else "-"
        pairs.append(("Usuário:", user_name))

        i = 0
        while i < len(row):
            val = row[i]
            if isinstance(val, str) and ":" in val and i + 1 < len(row):
                label = normalize_value(val)
                value_raw = row[i+1]
                value = normalize_value(value_raw).replace("\n", "").replace("\r", "").strip()
                if any(x.lower() in label.lower() for x in ["foto 1", "foto 2", "assinatura"]) and value.startswith("http"):
                    if "assinatura" in label.lower():
                        img_obj = fetch_image(value, max_w=IMG_MAX_W, max_h=SIGNATURE_MAX_H, align_center=True)
                    else:
                        img_obj = fetch_image(value, max_w=IMG_MAX_W, max_h=IMG_MAX_H)
                    pairs.append((label, img_obj if img_obj else value))
                else:
                    pairs.append((label, value if value else "-"))
                i += 2
            else:
                i += 1

        story = []
        reg_id = normalize_value(row[0]) if len(row) > 0 else f"registro_{r_index - 1}"
        pdf_file_name = f"{output_dir}/relatorio_{reg_id}.pdf"

        story.append(Paragraph(f"Registro {reg_id}", styles["ReportTitle"]))
        story.append(Spacer(1, 3))
        story.append(make_qa_table(pairs, 1, PAGE_W - (2 * MARGIN_SIDE_MM * mm)))
        doc = SimpleDocTemplate(pdf_file_name, pagesize=PAGE_SIZE, leftMargin=MARGIN_SIDE_MM*mm, rightMargin=MARGIN_SIDE_MM*mm)
        doc.build(story)
        generated_files.append(pdf_file_name)

    # Compacta os PDFs em um ZIP
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for f in generated_files:
            zipf.write(f, os.path.basename(f))
    zip_buffer.seek(0)
    return zip_buffer

# ===================== INTERFACE STREAMLIT =====================
st.set_page_config(page_title="Gerador de Relatórios Geodata", page_icon="📄", layout="centered")

st.title("📄 Gerador de Relatórios - Geodata")
st.write("Envie o arquivo Excel no formato esperado e clique em **Gerar PDFs**.")

uploaded_file = st.file_uploader("Selecione o arquivo Excel", type=["xlsx"])
if uploaded_file:
    if st.button("Gerar PDFs"):
        with st.spinner("Gerando PDFs, aguarde..."):
            try:
                zip_file = gerar_pdfs(uploaded_file)
                st.success("PDFs gerados com sucesso!")
                st.download_button(
                    label="⬇️ Baixar arquivos ZIP",
                    data=zip_file,
                    file_name="relatorios_individuais.zip",
                    mime="application/zip"
                )
            except Exception as e:
                st.error(f"Ocorreu um erro: {e}")
