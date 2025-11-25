import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import requests
import math
import os
import zipfile
import re
import unicodedata
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
LOGO_URL = "https://app.agrireport.agr.br/customers/picture/16?time=1759945360"
LOGO_HEIGHT_MM = 15.0
IMG_MAX_W = 130 * mm
IMG_MAX_H = 70 * mm
SIGNATURE_MAX_H = 25 * mm

# Estilo dos blocos (produtos e grupos)
BLOCK_BG = colors.HexColor("#F8FAFC")
BLOCK_BORDER = colors.HexColor("#CBD5E1")
PRODUCT_SPACER = 4 * mm  # pequeno espaço entre blocos

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
styles.add(ParagraphStyle(
    name="LabelSmall", parent=styles["Normal"], fontName=BASE_FONT,
    fontSize=7.8, leading=9.6, textColor=colors.HexColor("#4B5563")
))
styles.add(ParagraphStyle(
    name="Value", parent=styles["Normal"], fontName=BASE_FONT,
    fontSize=9.0, leading=11, textColor=colors.HexColor("#111827")
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

def is_url(s): 
    return isinstance(s, str) and s.strip().lower().startswith("http")

def looks_like_label(s): 
    return isinstance(s, str) and (":" in s or "?" in s)

def normalize_value(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()

def strip_accents(text: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", text)
        if unicodedata.category(c) != "Mn"
    )

def canonical_key(label: str) -> str:
    if label is None:
        return ""
    s = str(label).strip().lower()
    s = s.replace("?", ":")
    s = s.replace(" :", ":")
    if s.endswith(":"):
        s = s[:-1]
    s = strip_accents(s)
    s = re.sub(r"\s+", " ", s)
    return s

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
    for _ in range(pairs_per_row): 
        header += ["Pergunta", "Resposta"]

    formatted_pairs = []
    for q, a in pairs:
        q_par = q if isinstance(q, (Paragraph, Image)) else Paragraph(str(q) if q is not None else "-", styles["Q"])
        if isinstance(a, (Paragraph, Image)):
            a_par = a
        else:
            a_txt = str(a) if a is not None and str(a).strip() != "" else "-"
            a_par = Paragraph(a_txt, styles["A"])
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

# --------- Produtos: regex + campos esperados ---------
PRODUCT_REGEX = re.compile(r"^produto\s*(\d+)\s*:\s*$", re.IGNORECASE)
EXPECTED_PRODUCT_FIELDS = [
    "Produto {n}:",
    "Lote:",
    "Dose (ml/100kg):",
    "Utilizado (ml total):",
]

def _is_blank_value(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if isinstance(v, Image):
        return False
    s = str(v).strip().lower()
    return s in ("", "-", "n/a", "na", "null", "none")

def extract_products_and_rest(pairs, max_products=11):
    products = []
    rest = []
    i = 0
    used_idx = set()
    total = len(pairs)

    while i < total:
        q, a = pairs[i]
        m = PRODUCT_REGEX.match(str(q).strip())
        if m and len(products) < max_products:
            n = int(m.group(1))
            expected_labels = [lbl.format(n=n) for lbl in EXPECTED_PRODUCT_FIELDS]
            collected = {lbl: None for lbl in expected_labels}
            collected[expected_labels[0]] = a
            used_idx.add(i)

            j = i + 1
            while j < total:
                qj, aj = pairs[j]
                if PRODUCT_REGEX.match(str(qj).strip()):
                    break
                if str(qj).strip() in expected_labels[1:]:
                    collected[str(qj).strip()] = aj
                    used_idx.add(j)
                j += 1

            has_any_answer = any(not _is_blank_value(collected[lbl]) for lbl in expected_labels)

            if has_any_answer:
                items = []
                for lbl in expected_labels:
                    val = collected[lbl]
                    lbl_par = Paragraph(lbl, styles["Q"])
                    if isinstance(val, Image):
                        items.append((lbl_par, val))
                    elif isinstance(val, Paragraph):
                        items.append((lbl_par, val))
                    else:
                        txt = (str(val).strip() if not _is_blank_value(val) else "-")
                        items.append((lbl_par, Paragraph(txt, styles["A"])))

                products.append({"n": n, "items": items})

            i = j
            continue
        i += 1

    for idx, (q, a) in enumerate(pairs):
        if idx in used_idx:
            continue
        rest.append((q, a))

    products.sort(key=lambda d: d["n"])
    return products, rest

def make_product_block_table(product_items, available_width):
    q_frac, a_frac = 0.35, 0.65
    widths = [available_width * q_frac, available_width * a_frac]

    table = Table(product_items, colWidths=widths, hAlign="LEFT")
    table.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 0.7, BLOCK_BORDER),
        ("INNERGRID", (0,0), (-1,-1), 0.25, BLOCK_BORDER),
        ("BACKGROUND", (0,0), (-1,-1), BLOCK_BG),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]))
    return table


# ===================== GRUPOS DE INFORMAÇÕES =====================
GROUPS = [
    ["Data", "Máquina TS", "Supervisor OTM"],
    ["Canal", "Cidade", "UF", "Consultor Responsável"],
    ["Empresa Contratante", "Produtor", "Telefone do Produtor", "Cidade", "UF"],
    ["Cultura", "Variedade", "Lote", "Empresa", "Tipo", "Peso Total"],
]

SYNONYMS = {
    "data": ["data", "data aplicação", "data do ts", "data da aplicação"],
    "maquina ts": ["maquina ts", "máquina ts", "maquina de ts"],
    "supervisor otm": ["supervisor otm", "supervisor", "otm supervisor"],

    "canal": ["canal"],
    "cidade": ["cidade", "municipio", "município"],
    "uf": ["uf", "estado", "sigla uf"],
    "consultor responsavel": ["consultor responsavel", "consultor responsável", "responsavel tecnico", "responsável técnico"],

    "empresa contratante": ["empresa contratante", "contratante"],
    "produtor": ["produtor", "nome do produtor"],
    "telefone do produtor": ["telefone do produtor", "telefone produtor", "telefone"],
    "cultura": ["cultura"],
    "variedade": ["variedade", "cultivar"],
    "lote": ["lote", "lote geral"],
    "empresa": ["empresa", "empresa ts", "empresa executora"],
    "tipo": ["tipo"],
    "peso total": ["peso total", "peso", "peso total (kg)"],
}

def canon_from_display(label: str) -> str:
    return canonical_key(label)

def key_matches(label_in_sheet: str, wanted_display_label: str) -> bool:
    key = canonical_key(label_in_sheet)
    target = canon_from_display(wanted_display_label)
    syns = SYNONYMS.get(target, [target])
    syns = [canonical_key(s) for s in syns]
    return key in syns

def build_lookup(pairs):
    index = {i: (q, a) for i, (q, a) in enumerate(pairs)}
    pockets = {}
    for i, (q, a) in index.items():
        k = canonical_key(str(q))
        pockets.setdefault(k, []).append(i)
    return index, pockets

def pop_first_matching(index, pockets, wanted_display_label):
    target = canon_from_display(wanted_display_label)
    syns = SYNONYMS.get(target, [target])
    syns = [canonical_key(s) for s in syns]
    for s in syns:
        if s in pockets and pockets[s]:
            idx = pockets[s].pop(0)
            q, a = index.pop(idx, (None, None))
            if not pockets[s]:
                pockets.pop(s, None)
            return q, a
    return None, None

def make_inline_group_block(labels_order, index, pockets, available_width):
    n = len(labels_order)
    if n == 0:
        return None, []

    original_pairs = []
    values = []
    for lbl in labels_order:
        q, a = pop_first_matching(index, pockets, lbl)
        original_pairs.append((q, a))
        values.append(a)

    col_w = available_width / n
    col_widths = [col_w for _ in range(n)]

    top_row = [Paragraph(lbl, styles["LabelSmall"]) for lbl in labels_order]

    bottom_row = []
    for a in values:
        if isinstance(a, (Paragraph, Image)):
            bottom_row.append(a)
        else:
            txt = str(a).strip() if a is not None and str(a).strip() != "" else "-"
            bottom_row.append(Paragraph(txt, styles["Value"]))

    tbl = Table([top_row, bottom_row], colWidths=col_widths, hAlign="LEFT")
    tbl.setStyle(TableStyle([
        ("BOX", (0,0), (-1,-1), 0.7, BLOCK_BORDER),
        ("INNERGRID", (0,0), (-1,-1), 0.25, BLOCK_BORDER),
        ("BACKGROUND", (0,0), (-1,-1), BLOCK_BG),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]))
    return tbl, original_pairs


# ===================== INTERFACE STREAMLIT =====================
st.set_page_config(page_title="Gerador de Relatórios OTM", layout="centered")
st.title("📄 Gerador de Relatórios PDF - OTM")
st.caption("Gera um PDF individual por registro da planilha Excel.")

sample_path = "/mnt/data/Questionario_Guia_de_TS_V4 (6).xlsx"

uploaded_file = st.file_uploader("Faça upload do arquivo Excel (.xlsx)", type=["xlsx"])

# Se não houver upload, tenta usar arquivo de exemplo
if uploaded_file is None and os.path.exists(sample_path):
    st.info("Nenhum arquivo enviado — usando arquivo de exemplo presente no ambiente.")
    df_raw = pd.read_excel(sample_path, header=None)
else:
    if uploaded_file:
        df_raw = pd.read_excel(uploaded_file, header=None)
    else:
        st.warning("Envie um arquivo .xlsx para continuar.")
        st.stop()

# ===================== PROCESSAMENTO =====================

# Primeira coluna = labels
labels = df_raw.iloc[:, 0].tolist()

# Demais colunas = registros
records = df_raw.iloc[:, 1:]

if records.empty:
    st.error("A planilha não possui registros nas colunas após a primeira.")
    st.stop()

st.success(f"{records.shape[1]} registros encontrados — prontos para gerar PDF!")

# ===================== GERAR PDF =====================

if st.button("Gerar PDFs"):
    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for col_index in range(records.shape[1]):
            col = records.iloc[:, col_index]

            pairs = []
            for q, a in zip(labels, col):
                if is_url(a):
                    img = fetch_image(a, max_w=IMG_MAX_W, max_h=IMG_MAX_H, align_center=True)
                    if img:
                        pairs.append((q, img))
                        continue
                pairs.append((q, normalize_value(a)))

            products, rest_pairs = extract_products_and_rest(pairs)

            index, pockets = build_lookup(rest_pairs)

            pdf_buffer = BytesIO()
            doc = SimpleDocTemplate(
                pdf_buffer,
                pagesize=A4,
                leftMargin=MARGIN_SIDE_MM * mm,
                rightMargin=MARGIN_SIDE_MM * mm,
                topMargin=MARGIN_TOP_MM * mm,
                bottomMargin=MARGIN_BOTTOM_MM * mm
            )

            story = []
            story.append(Paragraph(REPORT_TITLE, styles["ReportTitle"]))
            story.append(Spacer(1, 4*mm))

            available_width = PAGE_W - 2*(MARGIN_SIDE_MM*mm)

            # ----- Blocos de grupos -----
            for group_labels in GROUPS:
                tbl, used = make_inline_group_block(group_labels, index, pockets, available_width)
                if tbl:
                    story.append(tbl)
                    story.append(Spacer(1, 4*mm))

            # ----- Blocos de produtos -----
            if products:
                story.append(Paragraph("Produtos Utilizados", styles["SectionTitle"]))
                story.append(Spacer(1, 2*mm))

                for prod in products:
                    tbl_prod = make_product_block_table(prod["items"], available_width)
                    story.append(tbl_prod)
                    story.append(Spacer(1, PRODUCT_SPACER))

            # ----- Resto dos campos (Q&A comuns) -----
            leftover_pairs = list(index.values())

            if leftover_pairs:
                story.append(Paragraph("Informações Complementares", styles["SectionTitle"]))
                story.append(Spacer(1, 2*mm))

                tbl = make_qa_table(leftover_pairs, pairs_per_row=1, available_width=available_width)
                story.append(tbl)

            # Salvar o PDF
            doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)

            filename = f"registro_{col_index+1}.pdf"
            zipf.writestr(filename, pdf_buffer.getvalue())

    st.download_button(
        "Baixar ZIP com PDFs",
        data=zip_buffer.getvalue(),
        file_name="relatorios_otm.zip",
        mime="application/zip"
    )

    st.success("Arquivos gerados com sucesso!")
