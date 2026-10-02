import streamlit as st
import pandas as pd
from pypdf import PdfReader
import io

st.set_page_config(page_title="PDF to Excel", layout="centered")
st.title("تبدیل PDF به Excel")

uploaded = st.file_uploader("فایل PDF را انتخاب کنید", type="pdf")

if uploaded:
    reader = PdfReader(uploaded)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    
    st.subheader("متن استخراج‌شده:")
    st.text_area("", text, height=300)
    
    df = pd.DataFrame({"متن": text.split("\n")})
    output = io.BytesIO()
    df.to_excel(output, index=False)
    
    st.download_button(
        label="دانلود Excel",
        data=output.getvalue(),
        file_name="output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
