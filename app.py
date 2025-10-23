import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.lib import colors

st.set_page_config(page_title="Gerar PDF do Questionário", layout="centered")

st.title("📋 Gerador de Relatório do Questionário")

uploaded_file = st.file_uploader("Envie o arquivo Excel", type=["xlsx"])

if uploaded_file:
    # Lê o Excel ignorando as duas primeiras linhas
    df = pd.read_excel(uploaded_file, header=None, skiprows=2)

    st.success(f"Arquivo carregado com {df.shape[0]} respostas e {df.shape[1]} colunas.")

    # Gera o PDF
    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    for i, row in df.iterrows():
        y = height - 2*cm
        pdf.setFont("Helvetica-Bold", 14)
        pdf.drawString(2*cm, y, f"Questionário #{i+1}")
        y -= 1*cm
        pdf.setFont("Helvetica", 10)

        # percorre de 2 em 2 (pergunta / resposta)
        for col in range(0, len(row), 2):
            pergunta = str(row[col]) if pd.notna(row[col]) else ""
            resposta = ""

            # evita erro se não houver a coluna de resposta correspondente
            if col + 1 < len(row):
                resposta = str(row[col + 1]) if pd.notna(row[col + 1]) else ""

            # ignora células totalmente vazias
            if not pergunta.strip() and not resposta.strip():
                continue

            # quebra página se o espaço estiver acabando
            if y < 2*cm:
                pdf.showPage()
                y = height - 2*cm
                pdf.setFont("Helvetica", 10)

            # pergunta
            pdf.setFillColor(colors.black)
            pdf.setFont("Helvetica-Bold", 9)
            pdf.drawString(2*cm, y, f"{pergunta}:")
            y -= 0.4*cm

            # resposta
            pdf.setFont("Helvetica", 9)
            pdf.setFillColor(colors.darkgray)
            text = pdf.beginText(2.5*cm, y)
            for line in resposta.splitlines():
                text.textLine(line)
                y -= 0.4*cm
            pdf.drawText(text)

            y -= 0.3*cm

        pdf.showPage()

    pdf.save()
    buffer.seek(0)

    st.download_button(
        label="📄 Baixar PDF Gerado",
        data=buffer,
        file_name="questionarios_geodata.pdf",
        mime="application/pdf"
    )

    st.success("PDF gerado com sucesso, incluindo todas as perguntas e respostas (inclusive DZ e EA).")
