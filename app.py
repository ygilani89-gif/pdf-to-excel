import streamlit as st
import pandas as pd
from pypdf import PdfReader
import io, re

st.set_page_config(page_title="تبدیل PDF به Excel", layout="centered")
st.title("تبدیل صورت مالی PDF به Excel")

uploaded = st.file_uploader("فایل PDF را انتخاب کنید", type="pdf")

if uploaded:
    reader = PdfReader(uploaded)
    rows = []
    for page in reader.pages:
        text = page.extract_text() or ""
        for line in text.split("\n"):
            line = line.strip()
            if not line:
                continue
            # جدا کردن عدد از متن
            match = re.match(r"^(.*?)[\s:]*([\d,،]+)\s*$", line)
            if match:
                desc = match.group(1).strip()
                amount = match.group(2).replace("،", ",").strip()
                rows.append({"شرح": desc, "مبلغ": amount})
            else:
                rows.append({"شرح": line, "مبلغ": ""})

    df = pd.DataFrame(rows)
    st.subheader("پیش‌نمایش:")
    st.dataframe(df, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="گزارش")
    st.download_button(
        label="⬇️ دانلود Excel",
        data=output.getvalue(),
        file_name="financial_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )