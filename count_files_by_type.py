import os
import csv
import sys
from datetime import datetime
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
    total_pages = {}
    total_rows = {}
    total_columns = {}

    for root, _, files in os.walk(directory):
        for file in files:
            ext = file.split('.')[-1].lower()
            file_path = os.path.join(root, file)
            file_size = os.path.getsize(file_path)

            try:
                modified_time = os.path.getmtime(file_path)
                modified_date = datetime.fromtimestamp(modified_time)
                year = modified_date.year
            except Exception as e:
                print(f"Error getting modified date for {file_path}: {e}")
                year = None

            if year is not None:
                file_counts.setdefault((ext, year), 0)
                file_sizes.setdefault((ext, year), 0)
                total_pages.setdefault((ext, year), 0)
                total_rows.setdefault((ext, year), 0)
                total_columns.setdefault((ext, year), 0)

                file_counts[(ext, year)] += 1
                file_sizes[(ext, year)] += file_size

                if ext == "pdf":
                    total_pages[(ext, year)] += count_pages_pdf(file_path)
                elif ext == "doc":
                    total_pages[(ext, year)] += count_pages_doc(file_path)
                elif ext == "docx":
                    total_pages[(ext, year)] += count_pages_docx(file_path)
                elif ext == "pptx":
                    total_pages[(ext, year)] += count_pages_pptx(file_path)
                elif ext == "xls":
                    rows, columns = count_rows_columns_xls(file_path)
                    total_rows[(ext, year)] += rows
                    total_columns[(ext, year)] += columns
                elif ext == "xlsx":
                    rows, columns = count_rows_columns_xlsx(file_path)
                    total_rows[(ext, year)] += rows
                    total_columns[(ext, year)] += columns

    rows = []
    for (ext, year), count in file_counts.items():
        row = [ext, year, count, file_sizes[(ext, year)], total_pages.get((ext, year), ""), total_rows.get((ext, year), ""), total_columns.get((ext, year), "")]
        rows.append(row)

    # Group rows by year and sort by count within each year (descending)
    rows_by_year = {}
    for row in rows:
        year = row[1]
        if year not in rows_by_year:
            rows_by_year[year] = []
        rows_by_year[year].append(row)
    
    for year in rows_by_year:
        rows_by_year[year].sort(key=lambda x: -x[2])  # Sort by count (descending)

    summary_totals = {}

    with open(output_file, mode='w', newline='') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(["File Type", "Year", "Count", "Total Size (bytes)", "Total Pages", "Total Rows", "Total Columns"])
        for year in sorted(rows_by_year.keys(), reverse=True):  # Sort years descending
            table = PrettyTable()
            table.field_names = ["File Type", "Year", "Count", "Total Size (bytes)", "Total Pages", "Total Rows", "Total Columns"]

            total_count = 0
            total_size = 0
            total_pages = 0
            total_rows = 0
            total_columns = 0

            for row in rows_by_year[year]:
                table.add_row(row)
                writer.writerow(row)
                total_count += row[2]
                total_size += row[3]
                total_pages += row[4] if row[4] else 0
                total_rows += row[5] if row[5] else 0
                total_columns += row[6] if row[6] else 0

                # Update summary totals
                ext = row[0]
                summary_totals.setdefault(ext, {"count": 0, "size": 0, "pages": 0, "rows": 0, "columns": 0})
                summary_totals[ext]["count"] += row[2]
                summary_totals[ext]["size"] += row[3]
                summary_totals[ext]["pages"] += row[4] if row[4] else 0
                summary_totals[ext]["rows"] += row[5] if row[5] else 0
                summary_totals[ext]["columns"] += row[6] if row[6] else 0

            totals_row = ["TOTALS", year, total_count, total_size, total_pages, total_rows, total_columns]
            table.add_row(totals_row)
            writer.writerow(totals_row)
            
            print(f"\nYear: {year}")
            print(table)
            writer.writerow([])  # Add a blank line between years for readability

    # Summary table
    summary_table = PrettyTable()
    summary_table.field_names = ["File Type", "Count", "Total Size (bytes)", "Total Pages", "Total Rows", "Total Columns"]

    total_summary_count = 0
    total_summary_size = 0
    total_summary_pages = 0
    total_summary_rows = 0
    total_summary_columns = 0

    with open(output_file, mode='a', newline='') as csv_file:  # Append to the CSV file
        writer = csv.writer(csv_file)
        writer.writerow([])
        writer.writerow(["Summary Totals"])
        writer.writerow(["File Type", "Count", "Total Size (bytes)", "Total Pages", "Total Rows", "Total Columns"])

        for ext, totals in summary_totals.items():
            row = [ext, totals["count"], totals["size"], totals["pages"], totals["rows"], totals["columns"]]
            summary_table.add_row(row)
            writer.writerow(row)
            total_summary_count += totals["count"]
            total_summary_size += totals["size"]
            total_summary_pages += totals["pages"]
            total_summary_rows += totals["rows"]
            total_summary_columns += totals["columns"]

        totals_row = ["TOTALS", total_summary_count, total_summary_size, total_summary_pages, total_summary_rows, total_summary_columns]
        summary_table.add_row(totals_row)
        writer.writerow(totals_row)

        print("\nSummary Totals")
        print(summary_table)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py <directory_to_search> [output_file]")
        sys.exit(1)

    directory_to_search = sys.argv[1]
    output_file_path = sys.argv[2] if len(sys.argv) > 2 else "output.csv"
    main(directory_to_search, output_file_path)
    