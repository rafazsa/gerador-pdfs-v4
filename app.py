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
SIGNATURE_MAX_H = 25 * mm

# ===================== ESTILOS =====================
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(
    name="ReportTitle", parent=styles["Heading1"], fontName=BASE_FONT,
    fontSize=12.5, leading=14, textColor=PRIMARY, spaceAfter=2, alignment=1
))
styles.add(ParagraphStyle(
    name="SectionTitle", parent=styles["Heading2"], fontName=BASE_FONT,
    fontSize=9.8, leading=12, textColor=PRIMARY, spaceBefore=2, spaceAfter=2
))
styles.add(ParagraphStyle(
    name="Q", parent=styles["Normal"], fontName=BASE_FONT,
    fontSize=8.6, leading=10.2, textColor=colors.HexColor("#111827")
))
styles.add(ParagraphStyle(
    name="A", parent=styles["Normal"], fontName=BASE_FONT,
    fontSize=8.6, leading=10.2, textColor=colors.HexColor("#111827")
))

# ===================== FUNÇÕES AUXILIARES =====================
def header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(ACCENT)
    canvas.rect(0, PAGE_H - 15*mm, PAGE_W, 15*mm, stroke=0, fill=1)
    try:
        response = requests.get(LOGO_URL, timeout=10)
        img_data = BytesIO(response.content)
        img = Image(img_data)
        img_w, img_h = img.wrap(0, 0)
        aspect = img_h / img_w
        logo_w = LOGO_HEIGHT_MM / aspect
        img._restrictSize(logo_w*mm, LOGO_HEIGHT_MM*mm)
        img.drawOn(canvas, MARGIN_SIDE_MM*mm, PAGE_H - 20*mm + (15*mm - LOGO_HEIGHT_MM)/2)
    except Exception:
        pass
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    canvas.setFont(BASE_FONT, 8)
    canvas.drawRightString(PAGE_W - MARGIN_SIDE_MM*mm, PAGE_H - 9*mm, f"Gerado em {now}")
    canvas.restoreState()

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

def is_url(s): return isinstance(s, str) and s.strip().lower().startswith("http")
def looks_like_label(s): return isinstance(s, str) and ":" in s
def normalize_value(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def pack_pairs_into_rows(pairs, pairs_per_row):
    rows, line = [], []
    for i, (q, a) in enumerate(pairs, 1):
        line.extend([q, a])
        if i % pairs_per_row == 0:
            rows.append(line); line = []
    if line:
        falta = (pairs_per_row*2) - len(line)
        line.extend([""] * falta)
        rows.append(line)
    return rows

def make_qa_table(pairs, pairs_per_row, available_width):
    q_frac, a_frac = 0.35, 0.65
    pair_unit = q_frac + a_frac
    widths = []
    for _ in range(pairs_per_row):
        widths += [available_width * (q_frac/pair_unit) / pairs_per_row,
                   available_width * (a_frac/pair_unit) / pairs_per_row]
    header = []
    for _ in range(pairs_per_row): header += ["Pergunta", "Resposta"]

    formatted_pairs = []
    for q, a in pairs:
        q_par = Paragraph(q, styles["Q"])
        a_par = a if isinstance(a, Image) else Paragraph(a if a else "-", styles["A"])
        formatted_pairs.append((q_par, a_par))

    data_rows = pack_pairs_into_rows(formatted_pairs, pairs_per_row)
    table = Table([header] + data_rows, colWidths=widths, hAlign="LEFT", repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), ACCENT),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), BASE_FONT),
        ("FONTSIZE", (0,0), (-1,0), 8.2),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.25, BORDER),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [ZEBRA_1, ZEBRA_2]),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))
    return table

# ===================== INTERFACE STREAMLIT =====================
st.set_page_config(page_title="Gerador de Relatórios Geodata", layout="centered")
st.title("📄 Gerador de Relatórios PDF - Geodata")
st.caption("Gera um PDF individual por registro da planilha Excel.")

uploaded_file = st.file_uploader("Faça upload do arquivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, header=None)
    if df_raw.shape[0] < 3:
        st.error("Planilha inesperada: preciso de pelo menos 3 linhas (título, cabeçalho e dados).")
        st.stop()

    st.success(f"✅ Arquivo carregado com {df_raw.shape[0]-2} registros.")

    if st.button("🚀 Gerar PDFs"):
        output_dir = "relatorios_individuais"
        os.makedirs(output_dir, exist_ok=True)
        generated_files = []

        progress = st.progress(0)
        total = df_raw.shape[0] - 2

        for idx, r_index in enumerate(range(2, df_raw.shape[0]), 1):
            row = df_raw.iloc[r_index].tolist()
            pairs = []

            user_name = normalize_value(row[2]) if len(row) > 2 else "-"
            pairs.append(("Usuário:", user_name))

            i = 0
            while i < len(row):
                val = row[i]
                if looks_like_label(val) and i + 1 < len(row):
                    label = normalize_value(val)
                    value_raw = row[i+1]
                    value = normalize_value(value_raw).replace("\n", "").replace("\r", "").strip()
                    is_image_field = any(x.lower() in label.lower() for x in [
                        "foto 1: semente tratada e não tratada",
                        "foto 2: embalagem dos produtos",
                        "assinatura do produtor ou responsável"
                    ])

                    if is_image_field and is_url(value):
                        if "assinatura" in label.lower():
                            img_obj = fetch_image(value, max_w=IMG_MAX_W, max_h=SIGNATURE_MAX_H, align_center=True)
                        else:
                            img_obj = fetch_image(value, max_w=IMG_MAX_W, max_h=IMG_MAX_H)
                        pairs.append((label, img_obj if img_obj else value))
                        i += 2
                        continue

                    pairs.append((label, value if value else "-"))
                    i += 2
                else:
                    i += 1

            reg_id = normalize_value(row[0]) if len(row) > 0 else f"registro_{r_index - 1}"
            pdf_file_name = f"{output_dir}/relatorio_{reg_id}.pdf"

            story = [Paragraph(f"Registro {reg_id}", styles["ReportTitle"]), Spacer(1, 3)]
            avail_w = PAGE_W - (2 * MARGIN_SIDE_MM * mm)
            qa_table = make_qa_table(pairs, 2, avail_w)
            story += [Paragraph("Perguntas e Respostas", styles["SectionTitle"]), qa_table, Spacer(1, 3)]

            doc = SimpleDocTemplate(
                pdf_file_name,
                pagesize=PAGE_SIZE,
                leftMargin=MARGIN_SIDE_MM*mm, rightMargin=MARGIN_SIDE_MM*mm,
                topMargin=MARGIN_TOP_MM*mm, bottomMargin=MARGIN_BOTTOM_MM*mm,
                title=REPORT_TITLE,
            )
            doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)
            generated_files.append(pdf_file_name)

            progress.progress(idx / total)

        if generated_files:
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for file_path in generated_files:
                    zf.write(file_path, os.path.basename(file_path))
            zip_buffer.seek(0)

            st.success(f"{len(generated_files)} PDFs gerados com sucesso!")
            st.download_button(
                label="📦 Baixar todos os PDFs (.zip)",
                data=zip_buffer,
                file_name="relatorios_individuais.zip",
                mime="application/zip"
            )
