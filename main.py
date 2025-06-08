
import streamlit as st
import tempfile
import os
from copy_summary_from_v3 import copy_updated_sheets_with_formatting

st.set_page_config(page_title="Excel AI Assistant", layout="centered")
st.title("📊 Excel AI Assistant")

st.markdown("""
Upload your Excel file and we'll automatically add:

- `Γενικό Αποτέλεσμα`
- `Διαφορά`

No manual editing needed — just upload and download.
""")

uploaded_file = st.file_uploader("📁 Upload your file", type=["xlsx"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_input,          tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_output:

        template_path = "bilio_with_v3_formulas.xlsx"  # Must be in the same repo
        temp_input.write(uploaded_file.read())

        copy_updated_sheets_with_formatting(
            template_path=template_path,
            target_path=temp_input.name,
            output_path=temp_output.name
        )

        original_filename = os.path.splitext(uploaded_file.name)[0]
        download_filename = f"{original_filename}_updated.xlsx"

        st.success("✅ Done! Click below to download your updated file.")

        with open(temp_output.name, "rb") as f:
            st.download_button(
                label="📥 Download updated file",
                data=f,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
