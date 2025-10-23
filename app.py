import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import requests
import math
import os
import zipfile
import re
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

PRODUCT_BLOCK_BG = colors.HexColor("#F8FAFC")  # fundo suave no bloco
PRODUCT_BLOCK_BORDER = colors.HexColor("#CBD5E1")  # borda suave
PRODUCT_SPACER = 4 * mm  # espaço pequeno entre blocos

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

# --------- NOVO: extração dos blocos de produtos ---------
PRODUCT_REGEX = re.compile(r"^produto\s*(\d+)\s*:\s*$", re.IGNORECASE)

EXPECTED_PRODUCT_FIELDS = [
    "Produto {n}:",
    "Lote:",
    "Dose (ml/100kg):",
    "Utilizado (ml total):",
]

def extract_products_and_rest(pairs, max_products=11):
    """
    Varre a lista de (label, value) e separa:
      - products: lista de dicts, um por produto {n, items=[(q,a),...]}
      - rest: lista de (q,a) que não pertencem aos blocos de produto
    É tolerante à ordem (pega campos do produto até antes do próximo produto ou fim).
    """
    products = []
    rest = []
    i = 0
    used_idx = set()
    # Mapear rapidamente para busca local
    total = len(pairs)

    while i < total:
        q, a = pairs[i]
        m = PRODUCT_REGEX.match(str(q).strip())
        if m and len(products) < max_products:
            n = int(m.group(1))
            # Inicializa o bloco com as 4 perguntas esperadas
            expected_labels = [lbl.format(n=n) for lbl in EXPECTED_PRODUCT_FIELDS]
            collected = {lbl: None for lbl in expected_labels}
            # O título "Produto n:" já temos
            collected[expected_labels[0]] = a
            used_idx.add(i)

            # Coleta sequencialmente até achar próximo produto ou acabar
            j = i + 1
            while j < total:
                qj, aj = pairs[j]
                if PRODUCT_REGEX.match(str(qj).strip()):
                    break  # próximo produto
                # Se for um dos campos esperados deste produto, coleta
                if str(qj).strip() in expected_labels[1:]:
                    collected[str(qj).strip()] = aj
                    used_idx.add(j)
                j += 1

            # Constrói a lista final de itens mantendo a ordem definida
            items = []
            for lbl in expected_labels:
                val = collected[lbl]
                lbl_par = Paragraph(lbl, styles["Q"])
                if isinstance(val, Image):
                    items.append((lbl_par, val))
                else:
                    items.append((lbl_par, Paragraph((val if val else "-"), styles["A"])))
            products.append({"n": n, "items": items})
            i = j
            continue
        i += 1

    # Tudo que não foi usado em blocos de produto vai para o restante
    for idx, (q, a) in enumerate(pairs):
        if idx in used_idx:
            continue
        q_par = Paragraph(q, styles["Q"])
        a_par = a if isinstance(a, Image) else Paragraph(a if a else "-", styles["A"])
        rest.append((q_par, a_par))

    # Ordena os produtos por número (caso apareçam fora de ordem)
    products.sort(key=lambda d: d["n"])
    return products, rest

def make_product_block_table(product_items, available_width):
    """
    Cria um bloco (tabela 2 col) para um único produto.
    Sem cabeçalho; 4 linhas: Produto n, Lote, Dose, Utilizado.
    """
    q_frac, a_frac = 0.35, 0.65
    widths = [available_width * q_frac, available_width * a_frac]
    table = Table(product_items, colWidths=widths, hAlign="LEFT")
    table.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 0.7, PRODUCT_BLOCK_BORDER),
        ("INNERGRID", (0,0), (-1,-1), 0.25, PRODUCT_BLOCK_BORDER),
        ("BACKGROUND", (0,0), (-1,-1), PRODUCT_BLOCK_BG),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
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

            # Exemplo de campo fixo (mantido)
            user_name = normalize_value(row[2]) if len(row) > 2 else "-"
            pairs.append(("Usuário:", user_name))

            # Varredura dinâmica de rótulos/valores
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

            # ===== NOVO: separar blocos de produto e demais =====
            products, rest_pairs = extract_products_and_rest(pairs, max_products=11)

            # ===== Montagem do PDF =====
            reg_id = normalize_value(row[0]) if len(row) > 0 else f"registro_{r_index - 1}"
            pdf_file_name = f"{output_dir}/relatorio_{reg_id}.pdf"

            story = [Paragraph(f"Registro {reg_id}", styles["ReportTitle"]), Spacer(1, 3)]
            avail_w = PAGE_W - (2 * MARGIN_SIDE_MM * mm)

            # Blocos de produto
            if products:
                story.append(Paragraph("Especificações dos Produtos", styles["SectionTitle"]))
                for p in products:
                    block_tbl = make_product_block_table(p["items"], avail_w)
                    story.append(block_tbl)
                    story.append(Spacer(1, PRODUCT_SPACER))  # pequeno espaço entre blocos

                # Pequeno espaço após o último produto antes das demais questões
                story.append(Spacer(1, PRODUCT_SPACER))

            # Demais questões
            if rest_pairs:
                story.append(Paragraph("Perguntas e Respostas (demais campos)", styles["SectionTitle"]))
                qa_table = make_qa_table(rest_pairs, 2, avail_w)
                story += [qa_table, Spacer(1, 3)]

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
