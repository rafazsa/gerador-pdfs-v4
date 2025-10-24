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
LOGO_URL = "https://report.geodata.com.br/customers/picture/16?time=1759945360"
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
    """Considera vazio quando No



