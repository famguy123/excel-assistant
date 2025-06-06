import streamlit as st
import tempfile
from copy_summary_διορθωμένο import copy_updated_sheets

st.set_page_config(page_title="Excel AI Assistant", layout="centered")
st.title("📊 AI Excel Assistant")

st.markdown("""
Upload your Excel file and we'll automatically add:

- `Γενικό Αποτέλεσμα`
- `Διαφορά`

...and fix any formula references.
""")

uploaded_file = st.file_uploader("📁 Upload your file", type=["xlsx"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_input, \
         tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_output:

        template_path = "bilio_with_v1_formulas.xlsx"  # This file stays in the repo
        temp_input.write(uploaded_file.read())

        copy_updated_sheets(
            template_path=template_path,
            target_path=temp_input.name,
            output_path=temp_output.name
        )

        st.success("✅ Done! Click below to download your updated file.")

        with open(temp_output.name, "rb") as f:
            st.download_button(
                label="📥 Download updated file",
                data=f,
                file_name="updated_file.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

