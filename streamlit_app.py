import streamlit as st
import os
import function
import tempfile

st.title("PDF → Word 自動轉換工具")

# 上傳 PDF
pdf_file = st.file_uploader("請上傳 PDF 檔案", type=["pdf"])

# 勾選要產出的格式
certif = st.checkbox("CB with Certif.")

# 勾選要產出的格式
gma_filter = st.checkbox("產出 GMA Word")
saa_filter = st.checkbox("產出 SAA Word")
stcoa_filter = st.checkbox("產出 STCOA Word")

if pdf_file and (gma_filter or saa_filter or stcoa_filter):
    with tempfile.TemporaryDirectory() as tmpdir:
        # 儲存 PDF
        pdf_path = os.path.join(tmpdir, pdf_file.name)
        with open(pdf_path, "wb") as f:
            f.write(pdf_file.read())

        download_buttons = []
        
        for name, check, template in [
            ("GMA", gma_filter, "format/GMA.docx"),
            ("SAA", saa_filter, "format/SAA.docx"),
            ("STCOA", stcoa_filter, "format/STCOA.docx")
        ]:
            if check:
                word_output_name = pdf_file.name.replace(".pdf", f"_{name}.docx")
                output_path = os.path.join(tmpdir, word_output_name)

                doc = function.run(pdf_path, template, certif)
                doc.save(output_path)

                with open(output_path, "rb") as out_file:
                    st.download_button(
                        label=f"下載 {name} Word 檔",
                        data=out_file,
                        file_name=word_output_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

