import streamlit as st
import pandas as pd
import io

def transform_pricing_data(input_df):
    """
    Transforms the input pricing data into the required structured format.
    """
    # Initialize output data list
    output_data = []

    # Extract deductible option names (column headers except 'Age')
    option_names = input_df.columns[1:]

    # Iterate through deductible options
    for option in option_names:
        for _, row in input_df.iterrows():
            age_str = str(row["Age"]).strip()

            # Determine AgeFrom and AgeTo
            if "-" in age_str:
                age_from, age_to = map(int, age_str.split("-"))
            else:
                age_from = age_to = int(age_str)

            # Extract premium value
            premium = row[option]

            # Skip rows where premium is missing or invalid
            if pd.isna(premium):
                continue

            # Determine InvoiceComponent:
            # - First 7 rows after the header should be "Member Dependent"
            # - Everything else should be "Member Premium"
            invoice_component = "Member Premium"
            if len(output_data) < 7:  # First 7 rows after header
                invoice_component = "Member Dependent"

            # Ensure independent members (18-59) are listed individually
            if invoice_component == "Member Premium" or (18 <= age_from <= 23):
                for age in range(age_from, age_to + 1):
                    output_data.append([
                        "", "", age, age, invoice_component, premium, premium, premium, option, "", ""
                    ])
            else:
                output_data.append([
                    "", "", age_from, age_to, invoice_component, premium, premium, premium, option, "", ""
                ])

        # Append an empty row to separate different deductible options
        output_data.append([""] * 11)

    # Convert to DataFrame
    output_columns = ["PlanCode", "RateZone", "AgeFrom", "AgeTo", "InvoiceComponent",
                      "Annual", "Renewal", "Transfer", "OptionName", "DateFrom", "DateTo"]
    df_output = pd.DataFrame(output_data, columns=output_columns)

    # Ensure AgeFrom and AgeTo have consistent formatting
    df_output["AgeFrom"] = pd.to_numeric(df_output["AgeFrom"], errors="coerce").astype("Int64")
    df_output["AgeTo"] = pd.to_numeric(df_output["AgeTo"], errors="coerce").astype("Int64")

    # Remove InvoiceComponent value in empty divider rows
    df_output.loc[df_output["AgeFrom"].isna(), "InvoiceComponent"] = ""

    return df_output

# Streamlit UI
st.title("Insurance Pricing Data Processor")

st.write("**Truly Prices Automation by Reinier**")

st.write("Upload an **Input.xlsx** file to generate the **Output.xlsx** file.")

# File upload
uploaded_file = st.file_uploader("Upload Input Excel File", type=["xlsx"])

if uploaded_file is not None:
    st.success("File uploaded successfully!")

    # Read the uploaded Excel file
    input_df = pd.read_excel(uploaded_file, sheet_name="Input")

    if st.button("Process Data"):
        with st.spinner("Processing..."):
            # Transform the data
            output_df = transform_pricing_data(input_df)

            # Save the output to an Excel file
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine="xlsxwriter") as writer:
                output_df.to_excel(writer, sheet_name="Output", index=False)
            output_buffer.seek(0)

            st.success("Processing complete! Download the Output.xlsx file below.")

            # Provide download link for the processed file
            st.download_button(label="Download Output.xlsx",
                               data=output_buffer,
                               file_name="Output.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
