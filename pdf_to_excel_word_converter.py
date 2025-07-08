import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
from docx import Document
import zipfile

st.set_page_config(page_title="PDF to Excel/Word Converter", layout="centered")

st.title("📄 PDF to Excel/Word Converter")

uploaded_file = st.file_uploader("Upload your PDF file", type=["pdf"])
output_format = st.radio("Select output format:", ("Excel", "Word"))
zip_option = st.checkbox("Compress output as .zip file")

def extract_text_lines(pdf_file):
    lines = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                page_lines = text.split('\n')
                lines.extend(page_lines)
    return lines

def convert_text_to_excel(lines):
    df = pd.DataFrame(lines, columns=["Text"])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted Text")
    output.seek(0)
    return output, "converted.xlsx"

def convert_text_to_word(lines):
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output, "converted.docx"

def create_zip(file_bytes, filename):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(filename, file_bytes.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

if uploaded_file:
    if st.button("Convert"):
        with st.spinner("Extracting and converting text..."):
            lines = extract_text_lines(uploaded_file)
            if output_format == "Excel":
                file_bytes, filename = convert_text_to_excel(lines)
            else:
                file_bytes, filename = convert_text_to_word(lines)

            if zip_option:
                zip_file = create_zip(file_bytes, filename)
                st.success("✅ Conversion complete! File compressed as ZIP.")
                st.download_button("Download ZIP File", zip_file, file_name="converted.zip")
            else:
                st.success(f"✅ Conversion to {output_format} complete!")
                st.download_button(f"Download {output_format} File", file_bytes, file_name=filename)
