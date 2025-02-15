import streamlit as st
import pandas as pd
import pytesseract
from pdf2image import convert_from_path
import os

# Set Tesseract path (Modify this based on your installation)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Function to extract text using OCR
def extract_text_with_ocr(pdf_path):
    transactions = []
    images = convert_from_path(pdf_path)

    for image in images:
        text = pytesseract.image_to_string(image)
        lines = text.split("\n")
        for line in lines:
            parts = line.split()
            if len(parts) >= 3:  # Assuming at least Date, Description, and Amount
                date, description, amount = parts[0], " ".join(parts[1:-1]), parts[-1]
                transactions.append([date, description, amount])

    return transactions

# Function to save to Excel
def save_to_excel(transactions, output_path):
    df = pd.DataFrame(transactions, columns=["Date", "Description", "Amount"])
    df.to_excel(output_path, index=False)

# Streamlit UI
st.title("📄 PDF to Excel Converter")
st.write("Upload a **bank statement PDF**, and it will be converted into an **Excel file**.")

# File uploader
uploaded_file = st.file_uploader("Choose a PDF file", type=["pdf"])

if uploaded_file:
    st.success("✅ File uploaded successfully!")
    
    # Save the uploaded file temporarily
    temp_pdf_path = "temp.pdf"
    with open(temp_pdf_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Convert PDF to Excel
    output_excel_path = "output.xlsx"
    transactions = extract_text_with_ocr(temp_pdf_path)

    if transactions:
        save_to_excel(transactions, output_excel_path)
        st.success("✅ Conversion successful!")

        # Provide a download button
        with open(output_excel_path, "rb") as f:
            st.download_button("📥 Download Excel File", f, file_name="converted.xlsx")
    else:
        st.error("⚠️ No transactions found in the PDF. Please check the file format.")
