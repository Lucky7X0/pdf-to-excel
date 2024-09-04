import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

def extract_table_data_from_text(text, current_date):
    data = []

    # Split the text into lines
    lines = text.split('\n')

    st.write("Text extracted from PDF:")
    st.write(lines)

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Check for a date at the beginning of the line and standardize it
        date_match = re.match(r'^\d{2}/\d{2}/\d{4}', line)
        if date_match:
            current_date = date_match.group(0)
            # Standardize date format
            current_date = pd.to_datetime(current_date, format='%d/%m/%Y').strftime('%d/%m/%Y')
            continue

        if not current_date:
            continue

        # Extract data using regex patterns
        user_id_match = re.search(r'\b[A-Za-z0-9]{4,}\b', line)
        punch_time_match = re.search(r'\d{2}:\d{2}:\d{2}', line)
        io_type_match = re.search(r'\bIN\b|\bOUT\b', line)

        user_id = user_id_match.group(0).strip() if user_id_match else ''
        punch_time = punch_time_match.group(0).strip() if punch_time_match else ''
        io_type = io_type_match.group(0).strip() if io_type_match else ''

        # Extract the name directly after the user ID
        if user_id:
            name_start = line.find(user_id) + len(user_id)
            name_end = line.find(punch_time) if punch_time else len(line)
            name = line[name_start:name_end].strip()
            name = re.sub(r'\bIN\b|\bOUT\b', '', name).strip()

            # Append the extracted data to the list
            if punch_time:
                data.append([current_date, user_id, name, punch_time, io_type])

    return data, current_date

def pdf_to_excel(pdf_file):
    all_data = []
    current_date = None

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            st.write(f"Text from page {page.page_number}:")
            st.write(text)

            if text:
                page_data, current_date = extract_table_data_from_text(text, current_date)
                if page_data:
                    all_data.extend(page_data)
            else:
                st.write("No text extracted from page.")

    if all_data:
        result_df = pd.DataFrame(all_data, columns=['Date', 'User ID', 'Name', 'Punch Time', 'I/O Type'])
        return result_df
    else:
        return pd.DataFrame()

def process_data(df):
    # Clean up the 'Punch Time' column and handle invalid formats
    df['Punch Time'] = df['Punch Time'].apply(
        lambda x: x if pd.to_datetime(x, format='%H:%M:%S', errors='coerce') is not pd.NaT else ''
    )

    # Standardize the 'Date' column to the format 'DD/MM/YYYY'
    df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y', errors='coerce').dt.strftime('%d/%m/%Y')
    
    return df

# Streamlit app
st.title("PDF to Excel Converter")

uploaded_file = st.file_uploader("Upload a PDF file", type="pdf")

if uploaded_file:
    st.write("Processing your file...")

    result_df = pdf_to_excel(uploaded_file)

    if not result_df.empty:
        st.write("Data extracted successfully!")
        st.dataframe(result_df)

        cleaned_df = process_data(result_df)

        if cleaned_df.empty:
            st.write("No data found or the data could not be processed.")
        else:
            st.write("Processed Data:")
            st.dataframe(cleaned_df)

            # Save the DataFrame to Excel and provide download link
            excel_buffer = BytesIO()
            cleaned_df.to_excel(excel_buffer, index=False, engine='openpyxl')
            excel_buffer.seek(0)

            st.download_button(
                label="Download Excel file",
                data=excel_buffer,
                file_name="output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.write("No data found in the PDF.")
