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
    # Considera rótulos que contenham ':' ou '?' (pergunta ou rótulo tradicional)
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
    s = s.replace("?", ":")  # normaliza "Pergunta?" para "Pergunta:"
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
    """
    Cria tabela de Q&A. Aceita q/a como strings ou flowables (Paragraph/Image).
    Idempotente: se já for Paragraph/Image, usa direto.
    """
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
    """Considera vazio quando None, '', '-' (após strip), ou NaN."""
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if isinstance(v, Image):
        return False
    s = str(v).strip().lower()
    return s in ("", "-","n/a","na","null","none")

def extract_products_and_rest(pairs, max_products=11):
    """
    Separa blocos de produtos (Produto n:, Lote:, Dose..., Utilizado...) do restante.
    - products: lista de dicts {"n": int, "items": [(Q_flowable, A_flowable), ...]}
      **Somente produtos com AO MENOS UM valor não-vazio entram.**
    - rest: lista de (q, a) CRUS (string/flowable) não pertencentes aos blocos.
      Campos de produtos vazios também são removidos do 'rest' (não aparecem no PDF).
    """
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
                    break  # próximo produto
                if str(qj).strip() in expected_labels[1:]:
                    collected[str(qj).strip()] = aj
                    used_idx.add(j)
                j += 1

            # Checagem de "produto vazio": todos os 4 campos sem conteúdo real
            has_any_answer = any(not _is_blank_value(collected[lbl]) for lbl in expected_labels)
            if has_any_answer:
                # Monta os itens do bloco (flowables) na ordem esperada
                items = []
                for lbl in expected_labels:
                    val = collected[lbl]
                    lbl_par = Paragraph(lbl, styles["Q"])
                    if isinstance(val, Image):
                        items.append((lbl_par, val))
                    elif isinstance(val, Paragraph):
                        items.append((lbl_par, val))
                    else:
                        items.append((lbl_par, Paragraph((str(val).strip() if not _is_blank_value(val) else "-"), styles["A"])))
                products.append({"n": n, "items": items})
            # Se for vazio: não adiciona aos produtos e mantém os índices em used_idx
            # para não aparecer em "demais campos"
            i = j
            continue
        i += 1

    # Demais pares não usados nos blocos de produto — manter CRU
    for idx, (q, a) in enumerate(pairs):
        if idx in used_idx:
            continue
        rest.append((q, a))

    products.sort(key=lambda d: d["n"])
    return products, rest

def make_product_block_table(product_items, available_width):
    """
    Bloco (tabela 2 col) para um único produto.
    """
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

# --------- NOVO: grupos de informações gerais ---------
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

# Caminho do arquivo de exemplo que você enviou (usado apenas se não fizer upload)
sample_path = "/mnt/data/Questionario_Guia_de_TS_V4 (6).xlsx"

uploaded_file = st.file_uploader("Faça upload do arquivo Excel (.xlsx)", type=["xlsx"])

# Se não houver upload, tenta usar o arquivo de exemplo
if uploaded_file is None and os.path.exists(sample_path):
    st.info("Nenhum arquivo enviado — usando arquivo de exemplo presente no ambiente.")
    df_raw = pd.read_excel(sample_path, header=None)
else:
    if uploaded_file:
        df_raw = pd.read_excel(uploaded_file, header=None)
    else:
        st.info("Aguardando upload do Excel...")
        st.stop()

# Verificação básica
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

    # Percorre as linhas de dados (a partir da linha 2, como no seu código original)
    for idx, r_index in enumerate(range(2, df_raw.shape[0]), 1):
        row_full_series = df_raw.iloc[r_index]
        # Converte para lista para indexação mais direta
        row_full = row_full_series.tolist()

        # REG_ID e outros campos que existem antes da coluna J devem ser lidos da linha completa
        reg_id = normalize_value(row_full[4]) if len(row_full) > 0 else f"registro_{r_index - 1}"
        pdf_file_name = f"{output_dir}/relatorio_{reg_id}.pdf"

        pairs = []

        # Campo fixo de exemplo (mantendo posição original: coluna C -> índice 2)
        user_name = normalize_value(row_full[8]) if len(row_full) > 2 else "-"
        pairs.append(("Usuário:", user_name))

        # --- AQUI: lemos apenas a partir da coluna J (índice 9) para montar os pares Q/A ---
        row_data = row_full[9:] if len(row_full) > 9 else []

        # Varredura dinâmica de rótulos/valores em row_data
        i = 0
        while i < len(row_data):
            val = row_data[i]
            if looks_like_label(val) and i + 1 < len(row_data):
                label = normalize_value(val)
                value_raw = row_data[i+1]
                value = normalize_value(value_raw).replace("\n", "").replace("\r", "").strip()

                # Campos de imagem conhecidos
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

        # 1) Separar blocos de produto (1..11) e demais perguntas
        products, rest_pairs = extract_products_and_rest(pairs, max_products=11)

        # 2) Dentro de "demais perguntas", organizar os grupos pedidos
        index, pockets = build_lookup(rest_pairs)

        # ===== Montagem do PDF =====
        story = [Paragraph(f"Registro {reg_id}", styles["ReportTitle"]), Spacer(1, 3)]
        avail_w = PAGE_W - (2 * MARGIN_SIDE_MM * mm)

        # (A) Grupos visuais das "demais informações"
        story.append(Paragraph("Informações Gerais", styles["SectionTitle"]))
        for group in GROUPS:
            tbl, _ = make_inline_group_block(group, index, pockets, avail_w)
            if tbl:
                story.append(tbl)
                story.append(Spacer(1, PRODUCT_SPACER))

        # (B) O que sobrar das "demais informações" vai para Q&A padrão
        remaining_pairs = list(index.values())  # ainda crus
        if remaining_pairs:
            story.append(Paragraph("Detalhes", styles["SectionTitle"]))
            qa_table = make_qa_table(remaining_pairs, 2, avail_w)
            story += [qa_table, Spacer(1, 3)]

        # (C) Depois os blocos de produto (apenas os com resposta)
        if products:
            story.append(Paragraph("Especificações dos Produtos", styles["SectionTitle"]))
            for p in products:
                block_tbl = make_product_block_table(p["items"], avail_w)
                story.append(block_tbl)
                story.append(Spacer(1, PRODUCT_SPACER))
            story.append(Spacer(1, PRODUCT_SPACER))

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
    else:
        st.warning("Nenhum PDF foi gerado.")

