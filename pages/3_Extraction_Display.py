import streamlit as st
import pandas as pd

st.write(st.session_state["shared"])

excel_path = "temp_files/output/DDR_extraction_output.xlsx"
final_df = pd.read_excel(excel_path)

st.subheader("Extraction Output Display",divider=True)

# Display the DataFrame with highlighting empty cells
st.experimental_data_editor(final_df,height=900,num_rows="dynamic")