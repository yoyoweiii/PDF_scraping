import pdfplumber
import openpyxl
import os

def extract_tables_and_text(pdf_path):
    print(pdf_path," is processing.")
    content = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                print(f"Page {page_num + 1}:")
                content+=f"Page {page_num + 1}:\n"
                text = page.extract_text()
                print("Text Content:")
                content += "\nText Content:\n"
                print(text)
                content += text
                for table_num, table in enumerate(page.extract_tables()):
                    print(f"Table {table_num + 1}:---------------------------")
                    for r in table:
                        print(r)
        print(pdf_path, " is done.")
    except:
        print("not valid path")
#pdf_path = r"C:\Users\TSECDC01\Desktop\SDxResearch\Line Class\CA\IEM OA4.pdf"
excel_file = 'file_plant_folder-match.xlsx'
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active
#pdf_path = r"C:\Users\TSECDC01\Desktop\SDxResearch\Line Class\CA\EDC AAG.pdf"
pdf_path = "C1A1.pdf"
extract_tables_and_text(pdf_path)
#workbook.save(excel_file)