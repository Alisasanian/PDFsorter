import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import shutil
import fitz
import time
import cv2
import re
import csv
import json
import io
from pathlib import Path
from PIL import Image
from paddleocr import PaddleOCR
from DEdependencies import bcolors
from DEdependencies import printProgressBar
from DEdependencies import display_time

"""     #3
The cropped drawing number PDFs take a LOT OF DATA to store, pretty much 1 : 1 in terms of the pdf being cropped since the crop is non-destructive.
This is bad because the OCR will have to go through a bigger file making the OCR process overall slower, alongside doubling your required PDF storage allocated data for no reason whatsoever.
In order to fix these issues, this script will sequentially convert each page of the pdf into a JPEG image of the cropped section (You lose all XREF data and other stuff).
The output compresses the used data down by over 96% of the original PDF size. Output is easier to scan as well for OCR.
WARNING: This script deletes the PDFs it is caching so make sure you turn that feature off if you want to keep those pdfs (why though)
"""

# ------------------------- Custom Functions --------------------------

total_time_start = time.time()

def PPP(page, pdfdpi=72, imgdpi=300):
    """
    PDF_PAGE_TO_PIL
    Converts a PyMuPDF page to a PIL Image for OCR processing.
    """
    mat = fitz.Matrix(imgdpi/pdfdpi, imgdpi/pdfdpi)
    pix = page.get_pixmap(matrix=mat)
    img_data = pix.pil_tobytes(format="JPEG")
    img = Image.open(io.BytesIO(img_data))
    return img

# -------------------------------- END --------------------------------

# Initialize PaddleOCR with speed optimizations
print(f"{bcolors.OKCYAN}Initializing PaddleOCR...{bcolors.ENDC}")
ocr_engine = PaddleOCR(
    lang='en',
    use_textline_orientation=False,
    device='gpu',
    precision='fp16'                # half‑precision for speed
)

dirpath = Path(__file__).parent.as_posix()

input_directory = Path(f"{dirpath}/_workingdata_/_pdfcache_")     # _pdfcache_
output_directory = Path(f"{dirpath}/_workingdata_/_indexdataset_/")

output_directory.mkdir(parents=True, exist_ok=True)     # Create output directory if it doesn't exist

file_paths = [f for f in input_directory.iterdir() if f.suffix.lower() == '.pdf' and f.stem.endswith('-drawingnoimage')]     # Get all image PDF files (ending with -drawingnoimage.pdf)

dataset = []    # Initialize the dataset

drawing_no_pattern = re.compile(r'(?:DRAWING|RAWING|AWING)\s*NO[:.]\s*([A-Z]+[0-9Oo]?(?:\.[0-9Oo]+[A-Z]?)+[^\s]*)')     # Regex pattern to find drawing numbers after "DRAWING NO: " with pattern AG.021.02.15.2A, capturing from start till blankspace
drawing_no_to_no = re.compile(r'(?:DRAWING|RAWING|AWING)\s*NO[:.]\s*.*?([A-Z]+[0-9Oo]*(?:\.[0-9Oo]+[A-Z]?)+)')          # Regex pattern to find drawing numbers after "DRAWING NO: " with pattern AG01.021.0.13.2A, capturing from start till blankspace
no_pattern = re.compile(r'[A-Z]+[0-9]+\.[0-9]+[A-Z]?')                                                                  # Regex pattern to find only drawing numbers with pattern AG023.295A
sheet_no_pattern = re.compile(r'(?:SHEET|HEET|EET)\s*NO[:.]\s*.*?([A-Z]+\-[0-9Oo]+\.[0-9Oo]+[-_]?)')                     # Regex pattern to find sheet numbers after "SHEET NO. ag _sh2glkha sghlka" with pattern AG02.91-
QF_drawing_no_pattern = re.compile(r'(?:DRAWING|RAWING|AWING)\s*NO[:.]\s*.*?([A-Z]+[0-9Oo]+\-[0-9Oo]+[A-Z]?)')          # Regex pattern to find drawing numbers after "DRAWING NO: " with pattern AB01-02A

successcount = 0
totalcount = 0
total_pages_processed = 0
total_a_or_g_count = 0
total_drawing_number_count = 0

for file_path in file_paths:
    start_time = time.time()        # Iterating over the files

    # Per-file counters
    file_pages_processed = 0
    file_a_or_g_count = 0
    drawing_number_count = 0

    try:
        print("")
        print(f"{'-' * 25}{bcolors.UNDERLINE}Processing: {file_path.name}{bcolors.ENDC}{'-' * 25}")
        print("")

        original_pdf_name = file_path.stem.replace('-drawingnoimage', '')        # Get the original PDF name (remove -drawingnoimage suffix)
        
        doc = fitz.open(str(file_path))     # Open the image PDF

        page_count = doc.page_count
        
        if doc.page_count == 0:
            print(f"{bcolors.WARNING}Warning: {file_path.name} has no pages, skipping...{bcolors.ENDC}")
            doc.close()
            continue
        
        file_results = []
        
        for page_num in range(doc.page_count):          # Iterating over each page in the opened file
            page = doc[page_num]

            img = PPP(page, 72, 300)      # Convert pdf page to PIL image

            img.save(f"{dirpath}/_workingdata_/Page{page_num}.jpg")    # Save the image so it can be ingested into PaddleOCR
            
            # Perform OCR
            try:
                img_path = f"{dirpath}/_workingdata_/Page{page_num}.jpg"
                ocrimg = cv2.imread(img_path)
                
                result = ocr_engine.predict(ocrimg)                 # Perform OCR with PaddleOCR

                if result and len(result) > 0:
                    # The result is a list with one dictionary per image
                    ocr_data = result[0]                            # Get first (and only) image result
                    
                    rec_texts = ocr_data.get('rec_texts', [])       # Extract recognized texts
                    
                    ocr_text = ' '.join(rec_texts)                  # Join all text pieces into a single string
                    #print("")
                    #print(ocr_text)
                    #print("")

                    drawing_number = None

                    # Try drawing_no_pattern first (general DRAWING NO: pattern)    
                    match = drawing_no_pattern.search(ocr_text)
                    if match:
                        drawing_number = match.group(1)
                        #print(f"Found via drawing_no_pattern: {drawing_number}")
                    
                    # If not found, try drawing_no_to_no (handles complex cases like "DRAWING NO: 8 A7.01")
                    if not drawing_number:
                        match = drawing_no_to_no.search(ocr_text)
                        if match:
                            drawing_number = match.group(1).strip()
                            #print(f"Found via drawing_no_to_no: {drawing_number}")
                    
                    # If still not found, try no_pattern (standalone pattern like "A4.04")
                    if not drawing_number:
                        match = no_pattern.search(ocr_text)
                        if match:
                            drawing_number = match.group(0)
                            #print(f"Found via no_pattern: {drawing_number}")

                    if not drawing_number:
                        match = sheet_no_pattern.search(ocr_text)
                        if match:
                            drawing_number = match.group(1)
                            #print(f"Found via sheet_no_pattern: {drawing_number}")

                    if not drawing_number:
                        match = QF_drawing_no_pattern.search(ocr_text)
                        if match:
                            drawing_number = match.group(1)
                            #print(f"Found via QF_drawing_no_pattern: {drawing_number}")

                if drawing_number:
                    # Clean up common OCR errors
                    drawing_number = drawing_number.strip()
                    drawing_number = re.sub(r'[Oo]', '0', drawing_number)
                    #print("drawing Number being used: " + drawing_number)

                    is_a_or_g = drawing_number[0].upper() in ['A', 'G']     # Check if it's an A or G drawing

                    result = {
                            'pdf_name': original_pdf_name,
                            'page_number': page_num + 1,
                            'drawing_number': drawing_number,
                            'A_or_G': is_a_or_g
                            }

                    if is_a_or_g:
                        file_a_or_g_count += 1
                        total_a_or_g_count += 1
                    
                    drawing_number_count += 1
                    total_drawing_number_count += 1

                    dataset.append(result)
                    file_results.append(result)

                else:
                    # No drawing number found
                    result = {
                        'pdf_name': original_pdf_name,
                        'page_number': page_num + 1,
                        'drawing_number': None,
                        'A_or_G': False
                    }
                    dataset.append(result)
                    file_results.append(result)

                file_pages_processed += 1
                total_pages_processed += 1

                printProgressBar(file_pages_processed, doc.page_count)
                
            except Exception as ocr_error:
                print(f"{bcolors.FAIL}  OCR error on page {page_num + 1}: {str(ocr_error)}{bcolors.ENDC}")
                file_pages_processed += 1
                total_pages_processed += 1
            
            os.remove(f"{dirpath}/_workingdata_/Page{page_num}.jpg")
        doc.close()

        end_time = time.time()
        file_time = end_time - start_time

        print(f"{bcolors.OKCYAN}Processing time: {file_time:.2f} seconds [{display_time(file_time)}]{bcolors.ENDC}")
        print(f"Drawing numbers in the document: {drawing_number_count}")
        print(f"Found {file_a_or_g_count} A_or_G drawing numbers from {file_pages_processed} pages")
        print("")

        successcount += 1
        totalcount += 1

    except Exception as e:
        print("")
        print("-" * 75)
        print(f"{bcolors.FAIL}Error processing {file_path.name}: {str(e)}{bcolors.ENDC}")
        print("-" * 75)
        print("")
        
        if 'doc' in locals():
            doc.close()
        
        totalcount += 1
        continue
    
    if file_a_or_g_count > 0:
        csv_path = output_directory / f"{original_pdf_name}_drawing_numbers_dataset.csv"
        with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
            if file_results:  # Use file_results for per-file output
                fieldnames = ['pdf_name', 'page_number', 'drawing_number', 'A_or_G']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(file_results)
    else:
        csv_path = output_directory / f"{original_pdf_name}_drawing_numbers_dataset_unsorted.csv"
        with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
            if file_results:  # Use file_results for per-file output
                fieldnames = ['pdf_name', 'page_number', 'drawing_number', 'A_or_G']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(file_results)

# Create the combined_data subfolder first
combined_data_dir = output_directory / "combined_data"
combined_data_dir.mkdir(parents=True, exist_ok=True)  # This creates the directory if it doesn't exist

# Save as CSV
csv_path = combined_data_dir / "combined_drawing_numbers_dataset.csv"
with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
    if dataset:
        fieldnames = ['pdf_name', 'page_number', 'drawing_number', 'A_or_G']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(dataset)

try:
    shutil.rmtree(input_directory)
    print(f"{bcolors.OKGREEN}Successfully cleaned up temporary directory: {input_directory}{bcolors.ENDC}")
except PermissionError as e:
    print(f"{bcolors.WARNING}Warning: Could not remove temporary directory {input_directory}: {e}{bcolors.ENDC}")
except Exception as e:
    print(f"{bcolors.FAIL}Error removing temporary directory: {e}{bcolors.ENDC}")

total_time_end = time.time()
elapsed_total_time = total_time_end - total_time_start

print("")
print("=" * 75)
print(f"{bcolors.OKGREEN}OCR EXTRACTION COMPLETE{bcolors.ENDC}")
print(f"Processed {successcount} files out of {totalcount} in {elapsed_total_time:.2f} seconds [{display_time(elapsed_total_time)}]")
print(f"Total pages processed: {len(dataset)}")
print(f"Total A_or_G drawing numbers extracted: {total_a_or_g_count} | out of {total_drawing_number_count} drawing numbers")
print("=" * 75)