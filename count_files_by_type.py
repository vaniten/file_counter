import os
import sys
import csv
from collections import defaultdict
import PyPDF2
import docx
import pandas as pd
from pptx import Presentation

def usage():
    print("Usage: python count_files_by_type.py <directory_to_search> [output_file]")
    sys.exit(1)

def count_pdf_pages(file_path):
    try:
        with open(file_path, 'rb') as f:
            pdf = PyPDF2.PdfFileReader(f)
            return pdf.getNumPages()
    except Exception as e:
        print(f"Error reading PDF {file_path}: {e}")
        return 0

def count_doc_pages(file_path):
    try:
        import subprocess
        result = subprocess.run(['antiword', '-m', '8859-1', file_path], stdout=subprocess.PIPE)
        return len(result.stdout.decode('utf-8').splitlines())
    except Exception as e:
        print(f"Error reading DOC {file_path}: {e}")
        return 0

def count_docx_pages(file_path):
    try:
        doc = docx.Document(file_path)
        return len(doc.paragraphs)
    except Exception as e:
        print(f"Error reading DOCX {file_path}: {e}")
        return 0

def count_xlsx_rows_columns(file_path):
    try:
        df = pd.read_excel(file_path)
        return df.shape
    except Exception as e:
        print(f"Error reading XLSX {file_path}: {e}")
        return 0, 0

def count_xls_rows_columns(file_path):
    try:
        df = pd.read_excel(file_path, engine='xlrd')
        return df.shape
    except Exception as e:
        print(f"Error reading XLS {file_path}: {e}")
        return 0, 0

def count_ppt_slides(file_path):
    try:
        import subprocess
        result = subprocess.run(['pptinfo', file_path], stdout=subprocess.PIPE)
        for line in result.stdout.decode('utf-8').splitlines():
            if line.startswith("Pages:"):
                return int(line.split()[1])
    except Exception as e:
        print(f"Error reading PPT {file_path}: {e}")
    return 0

def count_pptx_slides(file_path):
    try:
        prs = Presentation(file_path)
        return len(prs.slides)
    except Exception as e:
        print(f"Error reading PPTX {file_path}: {e}")
        return 0

def main():
    if len(sys.argv) < 2:
        usage()

    target_dir = sys.argv[1]

    if not os.path.isdir(target_dir):
        print(f"Error: Directory '{target_dir}' does not exist.")
        sys.exit(1)

    if len(sys.argv) >= 3:
        output_file = sys.argv[2]
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_file = os.path.join(script_dir, 'output.csv')

    file_types = defaultdict(int)
    file_sizes = defaultdict(int)
    total_pages = 0
    total_rows = 0
    total_columns = 0
    total_slides = 0

    for root, _, files in os.walk(target_dir):
        for file in files:
            file_path = os.path.join(root, file)
            extension = os.path.splitext(file)[1][1:].lower()
            file_size = os.path.getsize(file_path)
            file_sizes[extension] += file_size

            if extension:
                file_types[extension] += 1
            else:
                file_types['no_extension'] += 1

            # Calculate total pages for PDF, DOC, DOCX
            if extension == 'pdf':
                total_pages += count_pdf_pages(file_path)
            elif extension == 'doc':
                total_pages += count_doc_pages(file_path)
            elif extension == 'docx':
                total_pages += count_docx_pages(file_path)

            # Calculate total rows and columns for XLS, XLSX
            elif extension == 'xls':
                rows, columns = count_xls_rows_columns(file_path)
                total_rows += rows
                total_columns += columns
            elif extension == 'xlsx':
                rows, columns = count_xlsx_rows_columns(file_path)
                total_rows += rows
                total_columns += columns

            # Calculate total slides for PPT, PPTX
            elif extension == 'ppt':
                total_slides += count_ppt_slides(file_path)
            elif extension == 'pptx':
                total_slides += count_pptx_slides(file_path)

    # Output the results to the CLI
    print("File Type Counts:")
    for ext, count in file_types.items():
        print(f"{ext}: {count}")

    print("Total File Sizes:")
    for ext, size in file_sizes.items():
        print(f"{ext}: {size} bytes")

    print(f"Total number of pages (pdf, doc, docx): {total_pages}")
    print(f"Total number of rows (xls, xlsx): {total_rows}")
    print(f"Total number of columns (xls, xlsx): {total_columns}")
    print(f"Total number of slides (ppt, pptx): {total_slides}")

    # Write the results to the CSV file
    with open(output_file, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["File Type", "Count", "Total Size (bytes)"])
        for ext, count in file_types.items():
            writer.writerow([ext, count, file_sizes[ext]])
        writer.writerow(["Total Pages (pdf, doc, docx)", total_pages])
        writer.writerow(["Total Rows (xls, xlsx)", total_rows])
        writer.writerow(["Total Columns (xls, xlsx)", total_columns])
        writer.writerow(["Total Slides (ppt, pptx)", total_slides])

    print(f"Results have been written to {output_file}")

if __name__ == "__main__":
    main()
