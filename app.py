import streamlit as st
from extractor_lite import process_pdf

st.set_page_config(page_title="PDF → Excel (Cloud Lite)", page_icon="💳", layout="centered")

st.markdown("""
# 💳 PDF → Excel (Cloud Lite)
**Fast, reliable on Streamlit Cloud — text-based PDFs only (no OCR).**

- Transactions (Date, Description, Debit, Credit, Balance)
- Raw_Extract (every line we saw)
- Issues (unparsed/ignored)
- Metadata

For **scanned PDFs**, run the full local app with OCR.
""")

with st.sidebar:
    st.header("⚙️ Options")
    custom_yaml = st.text_area("Optional: paste a YAML template", value="", height=180)

uploaded = st.file_uploader("Upload a bank statement PDF (text-based)", type=["pdf"])
convert = st.button("🚀 Convert to Excel", disabled=(uploaded is None))

if convert and uploaded is not None:
    with st.spinner("Processing your PDF..."):
        excel_bytes, report = process_pdf(uploaded.read(), custom_yaml if custom_yaml.strip() else None)

    if excel_bytes is None:
        st.error(report.get("error") or "No text found; likely a scanned PDF. Please run locally with OCR.", icon="⚠️")
    else:
        st.success("Done! Download your Excel below.", icon="✅")
        st.write(f"- Method: **{report['method']}**")
        if report["notes"]:
            st.write("- Notes:")
            for n in report["notes"]:
                st.write(f"  - {n}")
        st.write(f"- Parsed transactions: **{report['transactions']}**")
        st.write(f"- Unparsed lines: **{report['unparsed']}**")
        st.write(f"- Ignored lines: **{report['ignored']}**")
        st.download_button("⬇️ Download Excel", data=excel_bytes, file_name="Statement.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
