import streamlit as st
import pandas as pd
import re
import unicodedata
from io import BytesIO
from docx import Document

st.title("Normalizador Word y Excel")

# -------------------------
# NORMALIZAR TEXTO
# -------------------------
def normalizar(texto):
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize('NFKD', texto)
    texto = texto.encode('ascii', 'ignore').decode('utf-8')
    texto = re.sub(r"[ .\-\/]+", "_", texto)
    texto = re.sub(r"[^a-z0-9_]", "", texto)
    texto = re.sub(r"_+", "_", texto)
    return texto.strip("_")

# -------------------------
# FUNCION NORMALIZAR WORD
# -------------------------
def normalizar_word(doc):

    # 🔹 PARÁRAFOS
    for para in doc.paragraphs:
        variables = re.findall(r"{{(.*?)}}", para.text)
        for var in variables:
            nueva = normalizar(var)
            para.text = para.text.replace(
                "{{" + var + "}}",
                "{{" + nueva + "}}"
            )

    # 🔥 TABLAS (AQUÍ ESTABA TU PROBLEMA)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:

                    variables = re.findall(r"{{(.*?)}}", para.text)

                    for var in variables:
                        nueva = normalizar(var)
                        para.text = para.text.replace(
                            "{{" + var + "}}",
                            "{{" + nueva + "}}"
                        )

# -------------------------
# SUBIR ARCHIVOS
# -------------------------
word_file = st.file_uploader("📄 Subir Plantilla Word", type=["docx"])
excel_file = st.file_uploader("📊 Subir data Excel", type=["xlsx"])

# -------------------------
# PROCESAR
# -------------------------
if word_file and excel_file:

    # ✅ WORD
    doc = Document(word_file)

    normalizar_word(doc)  # 🔥 CLAVE

    word_buffer = BytesIO()
    doc.save(word_buffer)

    # ✅ EXCEL
    df = pd.read_excel(excel_file, dtype=str)

    df.columns = [normalizar(col) for col in df.columns]

    df = df.dropna(how="all")

    if "nro" in df.columns:
        df["nro"] = df["nro"].fillna("").astype(str).str.strip()
        df = df[df["nro"] != ""]
        df = df[df["nro"] != "0"]

    df = df.fillna("").astype(str)
    df = df.apply(lambda col: col.str.strip())

    st.dataframe(df.head())

    # export excel
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    # DESCARGA
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
