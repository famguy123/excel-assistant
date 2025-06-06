import streamlit as st
import tempfile
import os
from copy_summary_διορθωμένο import copy_updated_sheets

st.title("📊 AI Excel Assistant")

st.markdown("""
Upload your template and working file.  
The app will copy:
- `v1` → **Γενικό Αποτέλεσμα**
- `v2` → **Διαφορά**

...and fix formula references.
""")

template_file = st.file_uploader("📁 Upload TEMPLATE file (with v1 & v2)", type=["xlsx"])
target_file = st.file_uploader("📁 Upload CLIENT file", type=["xlsx"])

if template_file and target_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_template, \
         tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_target, \
         tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_output:

        tmp_template.write(template_file.read())
        tmp_target.write(target_file.read())

        # Run the copy function
        copy_updated_sheets(
            template_path=tmp_template.name,
            target_path=tmp_target.name,
            output_path=tmp_output.name
        )

        st.success("✅ Processing complete!")

        with open(tmp_output.name, "rb") as f:
            st.download_button(
                label="📥 Download updated file",
                data=f,
                file_name="client_file_updated.xlsx"
            )
