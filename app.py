import streamlit as st
import pandas as pd
import re
import unicodedata
from zipfile import ZipFile
from io import BytesIO

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
    # NORMALIZAR WORD
    # =====================
    with ZipFile(word_file, "r") as zin:

        xml = zin.read("word/document.xml").decode("utf-8")
        variables = re.findall(r"{{(.*?)}}", xml)

        for var in variables:
            nueva = normalizar(var)
            xml = xml.replace("{{" + var + "}}", "{{" + nueva + "}}")

        word_buffer = BytesIO()

        with ZipFile(word_buffer, "w") as zout:
            for item in zin.infolist():
                if item.filename != "word/document.xml":
                    zout.writestr(item, zin.read(item.filename))
                else:
                    zout.writestr(item, xml.encode("utf-8"))

    # =====================
    # NORMALIZAR EXCEL
    # =====================
    df = pd.read_excel(excel_file, dtype=str)

    # normalizar nombres columnas
    df.columns = [normalizar(col) for col in df.columns]

    # ✅ eliminar filas vacías
    df = df.dropna(how="all")

    # ✅ eliminar filas con "total"
    df = df[
        ~df.apply(
            lambda row: row.astype(str).str.lower().str.contains("total").any(),
            axis=1
        )
    ]

    # ✅ eliminar filas numéricas (sumatorias)
    df = df[
        ~df.apply(
            lambda row: all(
                re.fullmatch(r"[\d\.,]+", str(val).strip())
                for val in row if str(val).strip() != ""
            ),
            axis=1
        )
    ]

    # ✅ eliminar filas con claves vacías o 0
    if "poliza" in df.columns:
        df = df[df["poliza"].str.strip() != ""]
        df = df[df["poliza"] != "0"]

    if "nro" in df.columns:
        df = df[df["nro"].str.strip() != ""]
        df = df[df["nro"] != "0"]

    # ✅ 🔥 convertir TODO a texto (CLAVE)
    df = df.fillna("").astype(str)

    # ✅ limpiar espacios
    df = df.apply(lambda col: col.str.strip())

    # ✅ evitar valores "nan"
    df = df.replace(["nan", "None", "NaN"], "")

    # =====================
    # PREVIEW
    # =====================
    st.subheader("✅ Vista previa limpia")
    st.dataframe(df.tail())

    st.success(f"Filas finales: {len(df)}")

    # =====================
    # EXPORTAR EXCEL LIMPIO
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
        "📊 Descargar Excel Limpio (texto)",
        data=excel_buffer.getvalue(),
        file_name="excel_limpio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
