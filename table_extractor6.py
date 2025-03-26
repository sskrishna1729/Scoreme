import fitz  # PyMuPDF
import pandas as pd
from openpyxl import Workbook
import os
import re

def sanitize_text(text):
    text = re.sub(r'[^\x20-\x7E]', '', text)  # Remove non-printable characters
    text = text.replace('\t', ' ').replace('\n', ' ').strip()  # Clean tabs and new lines
    return text

def extract_tables_from_pdf(pdf_path):
    pdf_document = fitz.open(pdf_path)
    tables = []

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text = page.get_text("text")  # Use plain text extraction for better structure
        lines = text.split("\n")
        
        detected_table = []
        for line in lines:
            cleaned_line = sanitize_text(line)
            if is_table(cleaned_line):
                table_data = extract_table_data(cleaned_line)
                detected_table.append(table_data)
        
        if detected_table:
            tables.append(detected_table)
    
    return tables

def is_table(line):
    return len(line.split()) > 2  # Detects if a line has multiple words (heuristic for tables)

def extract_table_data(line):
    return line.split()  # Splitting based on spaces to retain structure

def save_tables_to_excel(tables, excel_path):
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        for i, table_list in enumerate(tables):
            for j, table in enumerate(table_list):
                df = pd.DataFrame(table)
                df.to_excel(writer, sheet_name=f'Page_{i+1}_Table_{j+1}', index=False, header=False)

def main():
    pdf_path = os.path.join("Scoreme", "test6 (1).pdf")
    excel_path = os.path.join("Scoreme", "output_excel6.xlsx")
    tables = extract_tables_from_pdf(pdf_path)
    save_tables_to_excel(tables, excel_path)
    print(f"Extraction complete. Tables saved in: {excel_path}")

if __name__ == "__main__":
    main()
