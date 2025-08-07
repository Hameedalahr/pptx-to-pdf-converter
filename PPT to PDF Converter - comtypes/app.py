import streamlit as st
import os
import comtypes.client
import tempfile

def convert_ppt_to_pdf(ppt_path, output_dir):
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    ppt = powerpoint.Presentations.Open(ppt_path)
    pdf_path = os.path.join(output_dir, os.path.basename(ppt_path).replace(".pptx", ".pdf"))
    ppt.SaveAs(pdf_path, 32)  # 32 is the value for PDF format
    ppt.Close()
    powerpoint.Quit()
    return pdf_path

st.title("PPT to PDF Converter")

uploaded_file = st.file_uploader("Upload your .pptx file", type=["pptx"])

if uploaded_file is not None:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(uploaded_file.read())
        ppt_path = tmp.name

    output_dir = tempfile.mkdtemp()
    pdf_path = convert_ppt_to_pdf(ppt_path, output_dir)

    with open(pdf_path, "rb") as f:
        st.download_button(
            label="Download PDF",
            data=f,
            file_name=os.path.basename(pdf_path),
            mime="application/pdf"
        )

    st.success("Conversion successful!")
