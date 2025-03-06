import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill

# Titel en introductie
st.title("SEO Planning Tool")
st.write("Upload een Excel-bestand en krijg een SEO-planning met analyse!")

# Stap 1: Upload het ruwe Excel-bestand
uploaded_file = st.file_uploader("Upload je Excel-bestand", type=["xlsx"])

if uploaded_file:
    raw_data_df = pd.read_excel(uploaded_file)
    df_raw = raw_data_df.copy()

    # Outputconfiguratie
    output_config = [
        {"output_header": "Pagina URL", "input_header": "Address", "required": True, "avg_check": False},
        {"output_header": "Status code", "input_header": "Status Code", "required": True, "avg_check": False},
        {"output_header": "Indexeerbaar?", "input_header": "Indexability", "required": True, "avg_check": False},
        {"output_header": "URL in index?", "input_header": "Coverage", "required": True, "avg_check": False},
        {"output_header": "Laatste keer gecrawld?", "input_header": "Days Since Last Crawled", "required": True, "avg_check": True},
        {"output_header": "GSC - CTR", "input_header": "CTR", "required": True, "avg_check": True},
        {"output_header": "GSC - vertoningen", "input_header": "Impressions", "required": True, "avg_check": True},
        {"output_header": "GSC - klikken", "input_header": "Clicks", "required": True, "avg_check": True},
    ]

    # Bouw de output-DataFrame
    output_df = pd.DataFrame()
    for config in output_config:
        header = config["output_header"]
        input_header = config["input_header"]
        if input_header in raw_data_df.columns:
            output_df[header] = raw_data_df[input_header]
        else:
            if config["required"]:
                st.error(f"Verplichte kolom '{input_header}' ontbreekt.")
                st.stop()
            else:
                output_df[header] = ""

    # Voeg gemiddelde-checks toe
    for config in output_config:
        if config.get("avg_check", False):
            original_col = config["output_header"]
            new_col_name = f"Gem? ({original_col})"
            numeric_vals = pd.to_numeric(output_df[original_col], errors='coerce')
            avg_val = numeric_vals.mean()
            check_values = numeric_vals.apply(lambda x: "Meer" if x > avg_val else "Minder")
            output_df[new_col_name] = check_values

    # Toon de output DataFrame in Streamlit
    st.write("### Verwerkte SEO Planning")
    st.dataframe(output_df)

    # Download de verwerkte data als Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        output_df.to_excel(writer, sheet_name="SEO planning", index=False)
        df_raw.to_excel(writer, sheet_name="Raw data", index=False)
    
    st.download_button(
        label="Download verwerkte Excel",
        data=output.getvalue(),
        file_name="SEO_planning_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
