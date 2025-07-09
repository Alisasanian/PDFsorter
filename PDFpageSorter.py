import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import time
import csv
import pyairtable
import re
import pymupdf
from dotenv import find_dotenv, load_dotenv
from pathlib import Path
from DEdependencies import bcolors
from DEdependencies import printProgressBar
from DEdependencies import display_time

# ---------------------- Directories ----------------------
dirpath = Path(__file__).parent.as_posix()

input_directory = Path(f"{dirpath}/PDFsToProcess")
index_directory = Path(f"{dirpath}/_workingdata_/_indexdataset_/combined_data")
output_directory = Path(f"{dirpath}/SortedPDFs")

output_directory.mkdir(parents=True, exist_ok=True)

# ---------------------- Functions -----------------------

def parse_airtable_url(tableurl):
    """
    Parse an Airtable URL to extract BASE_ID, TABLE_NAME, and VIEW_NAME
    
    Expected URL format: https://airtable.com/appXXXXXXXXX/tblXXXXXXXXX/viwXXXXXXXXXX?blocks=hide
    
    Returns:
        tuple: (BASE_ID, TABLE_NAME, VIEW_NAME)
    """
    # Split the URL by forward slashes
    urllist = tableurl.split("/")

    try:
        BASE_ID = urllist[3]                # Extract BASE_ID - should be at index 3 and start with "app"
        if not BASE_ID.startswith("app"):
            raise ValueError(f"BASE_ID should start with 'app', got: {BASE_ID}")
        
        TABLE_NAME = urllist[4]             # Extract TABLE_NAME - should be at index 4 and start with "tbl"
        if not TABLE_NAME.startswith("tbl"):
            raise ValueError(f"TABLE_NAME should start with 'tbl', got: {TABLE_NAME}")
        
        view_with_params = urllist[5]       # Extract VIEW_NAME - should be at index 5 and start with "viw"
        VIEW_NAME = view_with_params.split("?")[0]      # Need to handle query parameters (everything after "?")
        if not VIEW_NAME.startswith("viw"):
            raise ValueError(f"VIEW_NAME should start with 'viw', got: {VIEW_NAME}")
        
        return BASE_ID, TABLE_NAME, VIEW_NAME
        
    except IndexError:
        raise ValueError(f"{bcolors.FAIL}URL format is incorrect. Expected format: https://airtable.com/appXXX/tblXXX/viwXXX{bcolors.ENDC}")
    except Exception as e:
        raise ValueError(f"{bcolors.FAIL}Error parsing URL: {e}{bcolors.ENDC}")

# ----------------------- API stuff -----------------------

dotenv_path = find_dotenv()
load_dotenv(dotenv_path)

API_KEY = os.getenv("PersonalAccessToken")

api = pyairtable.Api(API_KEY)

# ----------------------- Variables -----------------------

successcount = 0
totalcount = 0
total_pages_processed = 0
total_time_start = time.time()

tableurl = "https://airtable.com/appMB5vAVmKqJRyCW/tblkRMcFH2m4itvNF/viwp9wguAdGAVDBxY?blocks=hide"

BASE_ID, TABLE_NAME, VIEW_NAME = parse_airtable_url(tableurl)
table = api.table(BASE_ID, TABLE_NAME)
sheet_list = table.all(view = VIEW_NAME)

# ------------------------ Sorting -----------------------

sorted_page_numbers = []
missing_drawings = []
found_drawings = []

print("-" * 75)
print("Sorting Input PDFs according to Airtable index")
print("")

# Loop through each drawing number from airtable in the desired order
for i in range(len(sheet_list)):
    row = sheet_list[i]         # The actual data is in the 'fields' key
    fields = row["fields"]
    drawingno = fields.get("Sheet Number")
    # print(f"Drawing Number: {drawingno}")

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

                printProgressBar(i, len(sheet_list))
                # print(f"  -> Found on page {page_number}")
                break  # Stop searching once we found the match
    
    if not drawing_found:
        missing_drawings.append(drawingno)
        # print(f"{bcolors.WARNING}  -> WARNING: Drawing {drawingno} not found in CSV{bcolors.ENDC}")

print("")
print(f"\n--- SORTING SUMMARY ---")
print(f"Total drawings from Airtable: {len(sheet_list)}")
print(f"Found in CSV: {len(found_drawings)}")
print(f"{bcolors.WARNING}Total number of missing drawings from input PDF(s): {len(missing_drawings)}{bcolors.ENDC}")

if missing_drawings:
    print(f"{bcolors.WARNING}Missing drawings: {', '.join(missing_drawings)}{bcolors.ENDC}")

# print(f"Page order for sorting: {sorted_page_numbers}")
print("")
print("-" * 75)

# ---------------- Generating Sorted PDF -----------------

input_pdf_files = list(input_directory.glob("*.pdf"))

print("-" * 75)

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