import streamlit as st
import pandas as pd
import re
import unicodedata
from io import BytesIO
from docx import Document

st.title("Normalizador Word y Excel")

def normalizar(texto):
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize('NFKD', texto)
    texto = texto.encode('ascii', 'ignore').decode('utf-8')
    texto = re.sub(r"[ .\-\/]+", "_", texto)
    texto = re.sub(r"[^a-z0-9_]", "", texto)
    texto = re.sub(r"_+", "_", texto)
    return texto.strip("_")

# ✅ NUEVA FUNCIÓN SEGURA
def procesar_parrafo(paragraph):

    full_text = "".join(run.text for run in paragraph.runs)

    if "{{" not in full_text:
        return

    variables = re.findall(r"{{(.*?)}}", full_text)

    for var in variables:
        nueva = normalizar(var)
        full_text = full_text.replace(
            "{{" + var + "}}",
            "{{" + nueva + "}}"
        )

    index = 0
    for run in paragraph.runs:
        length = len(run.text)
        run.text = full_text[index:index + length]
        index += length

def normalizar_word(doc):

    for para in doc.paragraphs:
        procesar_parrafo(para)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    procesar_parrafo(para)

# -------------------------
# SUBIR ARCHIVOS
# -------------------------
word_file = st.file_uploader("📄 Subir Plantilla Word", type=["docx"])
excel_file = st.file_uploader("📊 Subir data Excel", type=["xlsx"])

if word_file and excel_file:

    doc = Document(word_file)
    normalizar_word(doc)

    word_buffer = BytesIO()
    doc.save(word_buffer)

    df = pd.read_excel(excel_file, dtype=str)
    df.columns = [normalizar(col) for col in df.columns]
    df = df.dropna(how="all")

    if "nro" in df.columns:
        df["nro"] = df["nro"].fillna("").astype(str).str.strip()
        df = df[df["nro"] != ""]
        df = df[df["nro"] != "0"]

    df = df.fillna("").astype(str)
    df = df.apply(lambda col: col.str.strip())

    st.subheader("✅ Datos limpios")
    st.dataframe(df.head())

    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    st.download_button(
        "📄 Descargar Word Normalizado",
        data=word_buffer.getvalue(),
        file_name="word_normalizado.docx"
    )

    st.download_button(
        "📊 Descargar Excel Limpio",
        data=excel_buffer.getvalue(),
        file_name="excel_limpio.xlsx"
    )
