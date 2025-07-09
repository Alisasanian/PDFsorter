import os
import fitz
import time
import pytesseract
import re
import csv
import json
import io
from pathlib import Path
from PIL import Image
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

dirpath = Path(__file__).parent.as_posix()

input_directory = Path(f"{dirpath}/_workingdata_/_pdfcache_")     # _pdfcache_
output_directory = Path(f"{dirpath}/_workingdata_/_indexdataset_/")

output_directory.mkdir(parents=True, exist_ok=True)     # Create output directory if it doesn't exist

file_paths = [f for f in input_directory.iterdir() if f.suffix.lower() == '.pdf' and f.stem.endswith('-drawingnoimage')]     # Get all image PDF files (ending with -drawingnoimage.pdf)

dataset = []    # Initialize the dataset

A_or_G_drawing_no_pattern = re.compile(r'(?:DRAWING|RAWING|AWING|WING|ING|NG|G)\s*NO[:.]\s*([AG][^\s]*)')      # Regex pattern to find drawing numbers after "DRAWING NO: " that start with A or G
drawing_no_pattern = re.compile(r'(?:DRAWING|RAWING|AWING|WING|ING|NG|G)\s*NO[:.]\s*([A-Z][^\s]*)')            # Regex pattern to find drawing numbers after "DRAWING NO: "

successcount = 0
totalcount = 0
total_pages_processed = 0
total_a_or_g_count = 0

total_time_start = time.time()

for file_path in file_paths:
    # Iterating over the files
    start_time = time.time()

    # Per-file counters
    file_pages_processed = 0
    file_a_or_g_count = 0

    try:
        print("")
        print(f"{'-' * 25}{bcolors.UNDERLINE}Processing: {file_path.name}{bcolors.ENDC}{'-' * 25}")
        print("")

        original_pdf_name = file_path.stem.replace('-drawingnoimage', '')        # Get the original PDF name (remove -drawingnoimage suffix)
        
        doc = fitz.open(str(file_path))     # Open the image PDF
        
        if doc.page_count == 0:
            print(f"{bcolors.WARNING}Warning: {file_path.name} has no pages, skipping...{bcolors.ENDC}")
            doc.close()
            continue
        
        file_results = []
        
        for page_num in range(doc.page_count):
            # Iterating over each page in the opened file
            page = doc[page_num]

            img = PPP(page,72,300)      # Convert pdf page to PIL iamge

            # img.save(f"C:/Users/BSSandbox/Desktop/PDFsorter/_workingdata_/Page{page_num}.jpg")    # Check what tesseract is seeing if you want to
            
            # Perform OCR
            try:
                ocr_text = pytesseract.image_to_string(img, config='--psm 11 --oem 1')      # Use config for fine tuning the OCR

                # print(f"  Page {page_num + 1} OCR text: {ocr_text.strip()}")    # Debug: print OCR result
                
                matches = A_or_G_drawing_no_pattern.findall(ocr_text)      # Search for drawing number pattern

                if matches:
                    A_or_G_drawing_number = ' '.join(matches[0].split()).strip()        # Clean up the drawing number (remove extra spaces, normalize)
                    A_or_G_drawing_number = re.sub(r'^GO(\d)', r'G\1', A_or_G_drawing_number)       # Replaces GO with G
                    A_or_G_drawing_number = re.sub(r'^A0Q', r'A0', A_or_G_drawing_number)           # Replaces A0Q with A0

                    # print(f"drawing_number: {drawing_number}")    # Debug: print OCR drawing_number regex result
                    
                    result = {
                            'pdf_name': original_pdf_name,
                            'page_number': page_num + 1,
                            'drawing_number': A_or_G_drawing_number,
                            }
                    
                    if A_or_G_drawing_number[0].upper() in ['A', 'G']:
                        result['A_or_G'] = True
                        file_a_or_g_count += 1
                        total_a_or_g_count += 1
                    else:
                        result['A_or_G'] = False

                    dataset.append(result)
                    file_results.append(result)


                else:
                    drawing_no_patten_match = drawing_no_pattern.findall(ocr_text)
                    
                    if drawing_no_patten_match:
                        drawing_number = ' '.join(drawing_no_patten_match[0].split()).strip()

                        result = {
                            'pdf_name': original_pdf_name,
                            'page_number': page_num + 1,
                            'drawing_number': drawing_number,
                            'A_or_G' : False
                        }
                        dataset.append(result)
                        file_results.append(result)
                    else:
                        result = {
                                'pdf_name': original_pdf_name,
                                'page_number': page_num + 1,
                                'drawing_number': None,
                                'A_or_G' : False
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

        doc.close()

        end_time = time.time()
        file_time = end_time - start_time

        print(f"{bcolors.OKCYAN}Processing time: {file_time:.2f} seconds [{display_time(file_time)}]{bcolors.ENDC}")
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

total_time_end = time.time()
elapsed_total_time = total_time_end - total_time_start

print("")
print("=" * 75)
print(f"{bcolors.OKGREEN}OCR EXTRACTION COMPLETE{bcolors.ENDC}")
print(f"Processed {successcount} files out of {totalcount} in {elapsed_total_time:.2f} seconds [{display_time(elapsed_total_time)}]")
print(f"Total A_or_G drawing numbers extracted: {total_a_or_g_count} | out of {len(dataset)} pages")
print("=" * 75)