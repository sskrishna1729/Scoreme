import fitz  # PyMuPDF
import pandas as pd
from openpyxl import Workbook
import os

def extract_tables_from_pdf(pdf_path):
    pdf_document = fitz.open(pdf_path)
    tables = []

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text = page.get_text("blocks")
        
        detected_table = []
        for block in text:
            if is_table(block):
                table_data = extract_table_data(block)
                detected_table.append(table_data)
        
        if detected_table:
            tables.append(detected_table)
    
    return tables

def is_table(block):
    return len(block[4].split()) > 2  # Simple heuristic for detecting tables

def extract_table_data(block):
    lines = block[4].split("\n")
    table_data = []
    for line in lines:
        # Split the line into columns based on consistent spacing
        columns = [col.strip() for col in line.split() if col]
        if columns:
            table_data.append(columns)
    return table_data

def save_tables_to_excel(tables, excel_path):
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        for i, table_list in enumerate(tables):
            for j, table in enumerate(table_list):
                df = pd.DataFrame(table)
                df.to_excel(writer, sheet_name=f'Page_{i+1}_Table_{j+1}', index=False, header=False)

def main():
    pdf_path = os.path.join("Scoreme", "test3 (1).pdf")
    excel_path = os.path.join("Scoreme", "output_excel3.xlsx")
    tables = extract_tables_from_pdf(pdf_path)
    save_tables_to_excel(tables, excel_path)
    print(f"Extraction complete. Tables saved in: {excel_path}")

if __name__ == "__main__":
    main()