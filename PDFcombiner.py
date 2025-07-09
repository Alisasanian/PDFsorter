import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import fitz
import time
from pathlib import Path
from DEdependencies import bcolors
from DEdependencies import display_time
from DEdependencies import printProgressBar

dirpath = Path(__file__).parent.as_posix()

input_directory = Path(f"{dirpath}/PDFsToProcess")   # For iterating over all the files

input_directory.mkdir(parents=True, exist_ok=True)   # Create output directory if it doesn't exist

file_paths = [f for f in input_directory.iterdir() if f.suffix.lower() == '.pdf']   # Get all PDF files in the directory

successcount = 0
totalcount = 0
total_pages = 0

total_time_start = time.time()

# Create a new empty PDF for the combined result
combined_doc = fitz.open()

print(f"{bcolors.HEADER}Starting PDF combination process...{bcolors.ENDC}")
print(f"Found {len(file_paths)} PDF files to combine")
print("")

total_pdfs = len(file_paths)
loadcount = 0

for file_path in file_paths:
    loadcount += 1
    start_time = time.time()     # START TIME OF FILE PROCESSING
    try:
        # dashes = "-" * 25
        # print(f"{dashes}{bcolors.UNDERLINE}Processing: {file_path.stem}.pdf{bcolors.ENDC}{dashes}")

        doc = fitz.open(str(file_path))     # Open the current PDF
        
        if doc.page_count == 0:
            print(f"{bcolors.WARNING}Warning: {file_path.name} has no pages, skipping...{bcolors.ENDC}")
            doc.close()
            totalcount += 1
            continue

        combined_doc.insert_pdf(doc)        # Insert all pages from the current PDF into the combined PDF
        
        pages_added = doc.page_count
        total_pages += pages_added
        
        # Close the current document
        doc.close()

        end_time = time.time()
        file_time = end_time - start_time
        
        successcount += 1
        totalcount += 1

        printProgressBar(loadcount, total_pdfs)

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

# Save the combined PDF if we successfully processed any files
if successcount > 0:
    try:
        # Generate output filename with timestamp to avoid overwriting
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"combined.pdf"
        
        print("")
        print(f"{bcolors.OKBLUE}Saving combined PDF...{bcolors.ENDC}")
        
        # Save the combined PDF
        combined_doc.save(str(f"{input_directory}/{output_filename}"))
        combined_doc.close()
        
        print(f"{bcolors.OKGREEN}Combined PDF saved as: {output_filename}{bcolors.ENDC}")
        print(f"Total pages in combined PDF: {total_pages}")
        
    except Exception as e:
        print(f"{bcolors.FAIL}Error saving combined PDF: {str(e)}{bcolors.ENDC}")
        combined_doc.close()
else:
    print(f"{bcolors.WARNING}No PDFs were successfully processed. No output file created.{bcolors.ENDC}")
    combined_doc.close()

# Cleanup old broken files
for file_path in file_paths:
    if file_path.name != "combined.pdf":
        os.remove(file_path)

total_time_end = time.time()
elapsed_total_time = total_time_end - total_time_start

print("")
print("=" * 75)
print(f"{bcolors.OKGREEN}PDF COMBINATION COMPLETE{bcolors.ENDC}")
print(f"Successfully combined {successcount} PDF(s) out of {totalcount} PDF(s)")
print(f"Total pages combined: {total_pages}")
print(f"Total time: {elapsed_total_time:.2f} seconds [{display_time(elapsed_total_time)}]")
print("=" * 75)