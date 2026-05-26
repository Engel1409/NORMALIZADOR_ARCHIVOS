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
# SUBIR ARCHIVOS
# -------------------------
word_file = st.file_uploader("📄 Subir Plantilla Word", type=["docx"])
excel_file = st.file_uploader("📊 Subir data Excel", type=["xlsx"])

# -------------------------
# PROCESAR
# -------------------------
if word_file and excel_file:

    # =====================
    # NORMALIZAR WORD ✅ (SOLUCIÓN REAL)
    # =====================
    doc = Document(word_file)

    for para in doc.paragraphs:

        # detectar variables correctamente
        variables = re.findall(r"{{(.*?)}}", para.text)

        for var in variables:
            nueva = normalizar(var)

            para.text = para.text.replace(
                "{{" + var + "}}",
                "{{" + nueva + "}}"
            )

    # ✅ GUARDAR WORD
    word_buffer = BytesIO()
    doc.save(word_buffer)

    # =====================
    # LIMPIAR EXCEL ✅
    # =====================
    df = pd.read_excel(excel_file, dtype=str)

    # normalizar columnas
    df.columns = [normalizar(col) for col in df.columns]

    # eliminar filas vacías
    df = df.dropna(how="all")

    # ✅ regla clave: eliminar si nro vacío o 0
    if "nro" in df.columns:
        df["nro"] = df["nro"].fillna("").astype(str).str.strip()
        df = df[df["nro"] != ""]
        df = df[df["nro"] != "0"]

    # ✅ convertir TODO a texto
    df = df.fillna("").astype(str)

    # limpiar espacios
    df = df.apply(lambda col: col.str.strip())

    # evitar valores raros
    df = df.replace(["nan", "None", "NaN"], "")

    # =====================
    # PREVIEW
    # =====================
    st.subheader("✅ Datos limpios")
    st.dataframe(df.head())

    st.success(f"Filas finales: {len(df)}")

    # =====================
    # EXPORTAR EXCEL
    # =====================
    excel_buffer = BytesIO()

    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    # =====================
    # DESCARGAS
    # =====================
    st.download_button(
        "📄 Descargar Word Normalizado",
        data=word_buffer.getvalue(),
        file_name="word_normalizado.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    st.download_button(
        "📊 Descargar Excel Limpio",
        data=excel_buffer.getvalue(),
        file_name="excel_limpio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
