import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import shutil
import time
import csv
import re
import pymupdf
from pathlib import Path
from DEdependencies import bcolors
from DEdependencies import printProgressBar
from DEdependencies import display_time

# ---------------------- Directories ----------------------
dirpath = Path(__file__).parent.as_posix()

input_directory = Path(f"{dirpath}/PDFsToProcess")
index_directory = Path(f"{dirpath}/_workingdata_/_indexdataset_/combined_data")
output_directory = Path(f"{dirpath}/SortedPDFs")
csv_dir = Path(f"{dirpath}/ReferenceCSV")

output_directory.mkdir(parents=True, exist_ok=True)
csv_dir.mkdir(parents=True, exist_ok=True)

csv_path = [f for f in csv_dir.iterdir() if f.suffix.lower() == '.csv']

file_paths = [f for f in input_directory.iterdir() if f.suffix.lower() == '.pdf']

# ----------------------- Variables -----------------------

successcount = 0
totalcount = 0
total_pages_processed = 0
total_time_start = time.time()

csv_name = csv_path[0]
sheet_list = []

with open(csv_name, 'r', encoding='utf-8') as f:
    reader = csv.reader(f)
    sheet_list = list(reader)

unprocessed_sheet_list = []
for row in sheet_list:
    unprocessed_sheet_list.append(row[0])

processed_sheet_list = []

drawing_no_pattern = re.compile(r'([A-Z]+[0-9Oo]?(?:\.[0-9Oo]+[A-Z]?)+[^\s]*)')     # Regex pattern to find drawing numbers with pattern AG.021.02.15.2A, capturing from start till blankspace
drawing_no_to_no = re.compile(r'([A-Z]+[0-9Oo]*(?:\.[0-9Oo]+[A-Z]?)+)')             # Regex pattern to find drawing numbers with pattern AG01.021.0.13.2A, capturing from start till blankspace
no_pattern = re.compile(r'[A-Z]+[0-9]+\.[0-9]+[A-Z]?')                              # Regex pattern to find only drawing numbers with pattern AG023.295A
sheet_no_pattern = re.compile(r'([A-Z]+\-[0-9Oo]+\.[0-9Oo]+[-_]?)')                 # Regex pattern to find sheet numbers with pattern AG02.91-
QF_drawing_no_pattern = re.compile(r'([A-Z]+[0-9Oo]+\-[0-9Oo]+[A-Z]?)')             # Regex pattern to find drawing numbers with pattern AB01-02A

for maybe_drawing_no in unprocessed_sheet_list:
    drawing_number = None
    maybe_drawing_no = str(maybe_drawing_no)

    match = drawing_no_pattern.search(maybe_drawing_no)
    if match:
        drawing_number = match.group(1)
        #print(f"Found via drawing_no_pattern: {drawing_number}")
    
    # If not found, try drawing_no_to_no (handles complex cases like "DRAWING NO: 8 A7.01")
    if not drawing_number:
        match = drawing_no_to_no.search(maybe_drawing_no)
        if match:
            drawing_number = match.group(1).strip()
            #print(f"Found via drawing_no_to_no: {drawing_number}")
    
    # If still not found, try no_pattern (standalone pattern like "A4.04")
    if not drawing_number:
        match = no_pattern.search(maybe_drawing_no)
        if match:
            drawing_number = match.group(0)
            #print(f"Found via no_pattern: {drawing_number}")

    if not drawing_number:
        match = sheet_no_pattern.search(maybe_drawing_no)
        if match:
            drawing_number = match.group(1)
            #print(f"Found via sheet_no_pattern: {drawing_number}")

    if not drawing_number:
        match = QF_drawing_no_pattern.search(maybe_drawing_no)
        if match:
            drawing_number = match.group(1)
            #print(f"Found via QF_drawing_no_pattern: {drawing_number}")
    
    if drawing_number is not None:
        processed_sheet_list.append(drawing_number)

# ------------------------ Sorting -----------------------

sorted_page_numbers = []
missing_drawings = []
found_drawings = []

print("-" * 75)
print("Sorting Input PDF(s) according to csv index")
print("")


# Loop through each drawing number from airtable in the desired order
for i in range(len(processed_sheet_list)):
    drawingno = processed_sheet_list[i]         # The actual data is in the 'fields' key
    drawing_found = False       # Flag to track if we found this drawing in the CSV
    
    with open(f"{index_directory}/combined_drawing_numbers_dataset.csv", 'r', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['pdf_name', 'page_number', 'drawing_number', 'A_or_G']
        reader = csv.DictReader(csvfile, fieldnames=fieldnames)

        next(reader, None)      # Skip the header row if it exists

        # Look through each row in the CSV for matching drawing number
        for row in reader:
            page_number = row['page_number']
            scanned_drawing_number = row["drawing_number"]

            # Check if this CSV row matches our current airtable drawing number
            if scanned_drawing_number.strip() == drawingno.strip():
                # Found a match - add this page number to our sorted list
                sorted_page_numbers.append(int(page_number))
                found_drawings.append(drawingno)
                drawing_found = True

                printProgressBar(i, len(processed_sheet_list))
                # print(f"  -> Found on page {page_number}")
                break  # Stop searching once we found the match
    
    if not drawing_found:
        missing_drawings.append(drawingno)
        # print(f"{bcolors.WARNING}  -> WARNING: Drawing {drawingno} not found in CSV{bcolors.ENDC}")

print("")
print(f"\n--- SORTING SUMMARY ---")
print(f"Total drawings from master CSV: {len(processed_sheet_list)}")
print(f"Found in index CSV: {len(found_drawings)}")
print(f"{bcolors.WARNING}Missing from CSV: {len(missing_drawings)}{bcolors.ENDC}")

if missing_drawings:
    print(f"{bcolors.WARNING}Missing drawings: {', '.join(missing_drawings)}{bcolors.ENDC}")

# print(f"Page order for sorting: {sorted_page_numbers}")
print("")

# ---------------- Generating Sorted PDF -----------------

input_pdf_files = list(input_directory.glob("*.pdf"))

if not input_pdf_files:
    print(f"{bcolors.FAIL}ERROR: No PDF files found in input directory{bcolors.ENDC}")
else:
    for i in range(len(input_pdf_files)):
        # Process the first PDF file found
        input_pdf_path = input_pdf_files[i]
        # print(f"\nProcessing PDF: {input_pdf_path.name}")

        source_doc = pymupdf.open(input_pdf_path)
        sorted_doc = pymupdf.open()
        pages_added = 0

        print("")
        print("Saving Sorted PDF...")
        print("")
        
        for page_num in sorted_page_numbers:
            page_index = page_num - 1

            if page_index < source_doc.page_count:
                sorted_doc.insert_pdf(source_doc, from_page=page_index, to_page=page_index)
                pages_added += 1
                # print(f"Added page {page_num} to sorted PDF")
            else:
                print(f"{bcolors.WARNING}WARNING: Page {page_num} doesn't exist in source PDF (only has {source_doc.page_count} pages){bcolors.ENDC}")

    output_filename = f"SORTED_{input_pdf_path.stem}.pdf"
    output_path = output_directory / output_filename
    sorted_doc.save(output_path)
    print(f"{bcolors.OKGREEN}\nSorted PDF saved as: {output_filename}{bcolors.ENDC}")
    print(f"{bcolors.OKGREEN}Pages in sorted PDF: {pages_added}{bcolors.ENDC}")
    print("-" * 75)

    source_doc.close()
    sorted_doc.close()

    if pages_added > 0:
        successcount += 1
        total_pages_processed += pages_added
    
    totalcount += 1

working_data = f"{dirpath}/_workingdata_/"
reference_CSV = f"{dirpath}/ReferenceCSV/"
PDFsToProcess = f"{dirpath}/PDFsToProcess/"
shutil.rmtree(working_data)
shutil.rmtree(reference_CSV)
shutil.rmtree(PDFsToProcess)