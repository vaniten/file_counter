import os
import csv
import sys
from PyPDF2 import PdfReader
import docx
import xlrd
from pptx import Presentation
from openpyxl import load_workbook
from prettytable import PrettyTable

def count_pages_pdf(file_path):
    try:
        reader = PdfReader(file_path)
        return len(reader.pages)
    except Exception as e:
        print(f"Error reading PDF {file_path}: {e}")
        return 0

def count_pages_doc(file_path):
    try:
        with open(file_path, 'rb') as f:
            data = f.read()
            return data.count(b'\f') + 1
    except Exception as e:
        print(f"Error reading DOC {file_path}: {e}")
        return 0

def count_pages_docx(file_path):
    try:
        doc = docx.Document(file_path)
        return len(doc.element.xpath('//w:sectPr'))
    except Exception as e:
        print(f"Error reading DOCX {file_path}: {e}")
        return 0

def count_pages_pptx(file_path):
    try:
        prs = Presentation(file_path)
        return len(prs.slides)
    except Exception as e:
        print(f"Error reading PPTX {file_path}: {e}")
        return 0

def count_rows_columns_xls(file_path):
    try:
        workbook = xlrd.open_workbook(file_path)
        total_rows = 0
        total_columns = 0
        for sheet in workbook.sheets():
            total_rows += sheet.nrows
            total_columns += sheet.ncols
        return total_rows, total_columns
    except Exception as e:
        print(f"Error reading XLS {file_path}: {e}")
        return 0, 0

def count_rows_columns_xlsx(file_path):
    try:
        workbook = load_workbook(filename=file_path, read_only=True)
        total_rows = 0
        total_columns = 0
        for sheet in workbook:
            max_row = sheet.max_row or 0
            max_column = sheet.max_column or 0
            total_rows += max_row
            total_columns += max_column
        return total_rows, total_columns
    except Exception as e:
        print(f"Error reading XLSX {file_path}: {e}")
        return 0, 0

def main(directory, output_file="output.csv"):
    file_counts = {}
    file_sizes = {}
    total_pages = {"pdf": 0, "doc": 0, "docx": 0, "pptx": 0}
    total_rows = {"xls": 0, "xlsx": 0}
    total_columns = {"xls": 0, "xlsx": 0}

    for root, _, files in os.walk(directory):
        for file in files:
            ext = file.split('.')[-1].lower()
            file_path = os.path.join(root, file)
            file_size = os.path.getsize(file_path)
            file_counts[ext] = file_counts.get(ext, 0) + 1
            file_sizes[ext] = file_sizes.get(ext, 0) + file_size

            if ext == "pdf":
                total_pages["pdf"] += count_pages_pdf(file_path)
            elif ext == "doc":
                total_pages["doc"] += count_pages_doc(file_path)
            elif ext == "docx":
                total_pages["docx"] += count_pages_docx(file_path)
            elif ext == "pptx":
                total_pages["pptx"] += count_pages_pptx(file_path)
            elif ext == "xls":
                rows, columns = count_rows_columns_xls(file_path)
                total_rows["xls"] += rows
                total_columns["xls"] += columns
            elif ext == "xlsx":
                rows, columns = count_rows_columns_xlsx(file_path)
                total_rows["xlsx"] += rows
                total_columns["xlsx"] += columns

    rows = []
    for ext in file_counts:
        if ext in total_pages:
            rows.append([ext, file_counts[ext], file_sizes[ext], total_pages[ext], "", ""])
        elif ext in total_rows:
            rows.append([ext, file_counts[ext], file_sizes[ext], "", total_rows[ext], total_columns[ext]])
        else:
            rows.append([ext, file_counts[ext], file_sizes[ext], "", "", ""])

    # Sort rows by total pages (descending) and then alphabetically by file type
    rows.sort(key=lambda x: (-x[3] if x[3] else 0, x[0]))

    table = PrettyTable()
    table.field_names = ["File Type", "Count", "Total Size (bytes)", "Total Pages", "Total Rows", "Total Columns"]

    for row in rows:
        table.add_row(row)

    print(table)

    with open(output_file, mode='w', newline='') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["File Type", "Count", "Total Size (bytes)", "Total Pages", "Total Rows", "Total Columns"])

        for row in rows:
            writer.writerow(row)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py <directory_to_search> [output_file]")
        sys.exit(1)

    directory_to_search = sys.argv[1]
    output_file_path = sys.argv[2] if len(sys.argv) > 2 else "output.csv"
    main(directory_to_search, output_file_path)
