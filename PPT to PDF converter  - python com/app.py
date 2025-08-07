import streamlit as st
import os
import pythoncom
from win32com.client import Dispatch


st.title(" PPTX to PDF Converter")

uploaded_file = st.file_uploader("Upload your PPTX file", type=["pptx"])

if uploaded_file:
    with open(uploaded_file.name, "wb") as f:
        f.write(uploaded_file.getbuffer())

    pptx_path = os.path.abspath(uploaded_file.name)
    pdf_path = pptx_path.replace(".pptx", ".pdf")

    try:
        pythoncom.CoInitialize()

        powerpoint = Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1
        presentation = powerpoint.Presentations.Open(pptx_path)
        presentation.SaveAs(pdf_path, 32) 
        presentation.Close()
        powerpoint.Quit()

        st.success("Conversion successful!")
        with open(pdf_path, "rb") as f:
            st.download_button(" Download PDF", f, file_name=os.path.basename(pdf_path))

    except Exception as e:
        st.error(f" Failed to convert: {e}")
