import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import fitz
import time
import io
from pathlib import Path
from PIL import Image
from DEdependencies import bcolors
from DEdependencies import get_folder_size_os
from DEdependencies import format_bytes
from DEdependencies import printProgressBar
from DEdependencies import display_time

"""     #2
This script takes PDFs containing the DRAWING NO: cropped out (through lossless means) and converts them into space-efficient pdfs
that purge all XREF data, only leaving behind small lossy JPEG images of each cropped page's content which is ready to be OCRed

Feel free to switch around the quality settings if you're having trouble with OCR image quality imput. I should have an envrionment file

with some kind of terminal-based GUI that handles this.
"""

# ------------------------- Custom Functions --------------------------

def PPP(page, pdfdpi = 72, imgdpi = 300):
    """
    PDF_PAGE_TO_PIL

    Converts a PyMuPDF page to a PIL Image for OCR processing. Handles in JPEG (No hate to my PNG fans)
    
    Args:
        page (fitz.Page): PyMuPDF page object
        pdfdpi (int): DPI of the pdf you want to convert from (Defualt is 72 for PyMuPDFs but change however you want)
        imgdpi (int): DPI of the image you want as output (Default is 300 which is really high quality but JPEG saves on storage)
        
    Returns:
        PIL.Image: PIL Image object ready for OCR processing
    """
    # Create transformation matrix for the specified DPI
    mat = fitz.Matrix(imgdpi/pdfdpi, imgdpi/pdfdpi)
    
    # Convert page to pixmap (raster image)
    pix = page.get_pixmap(matrix=mat)
    
    # Convert pixmap to PIL Image
    img_data = pix.pil_tobytes(format="JPEG")
    img = Image.open(io.BytesIO(img_data))
    
    return img

# -------------------------------- END --------------------------------

dirpath = Path(__file__).parent.as_posix()

input_directory = Path(f"{dirpath}/_workingdata_/_bloatedcache_")       # For iterating over all the files
output_directory = Path(f"{dirpath}/_workingdata_/_pdfcache_/")      # Put the finished files here

output_directory.mkdir(parents=True, exist_ok=True)     # Create output directory if it doesn't exist

file_paths = [f for f in input_directory.iterdir() if f.suffix.lower() == '.pdf' and f.stem.endswith('-drawingno')]     # Get all cropped PDF files (ending with -drawingno.pdf)

input_directory_size = get_folder_size_os(input_directory)

successcount = 0
totalcount = 0

total_time_start = time.time()

for file_path in file_paths:
    start_time = time.time()
    try:
        print("")
        print(f"{'-' * 25}{bcolors.UNDERLINE}Processing: {file_path.name}{bcolors.ENDC}{'-' * 25}")
        print("")

        doc = fitz.open(str(file_path))     # Open the cropped PDF
        
        if doc.page_count == 0:
            print(f"{bcolors.WARNING}Warning: {file_path.name} has no pages, skipping...{bcolors.ENDC}")
            doc.close()
            continue

        image_pdf = fitz.open()     # Create a new PDF to store images

        for page_num in range(doc.page_count):
            page = doc[page_num]
            
            img = PPP(page,72,300)
            
            img_width, img_height = img.size    # Create a new page in the output PDF with the same dimensions as the image

            page_width = img_width * 72 / 300    # Convert pixels to points
            page_height = img_height * 72 / 300  # Convert pixels to points
            
            new_page = image_pdf.new_page(width=page_width, height=page_height)
            
            # Insert the image into the new page
            img_bytes = io.BytesIO()
            img.save(img_bytes, format='JPEG')
            img_bytes.seek(0)
            
            # Insert image to fill the entire page
            rect = fitz.Rect(0, 0, page_width, page_height)
            new_page.insert_image(rect, stream=img_bytes)

            printProgressBar(page_num,doc.page_count)       # Live progress update in terminal
        


        original_name = file_path.stem.replace('-drawingno', '')        # Remove the -drawingno suffix
        output_filename = f"{original_name}-drawingnoimage.pdf"         # add -drawingnoimage
        output_path = output_directory / output_filename
        
        image_pdf.save(str(output_path))        # Save the new PDF with images
        image_pdf.close()
        doc.close()

        os.remove(str(file_path))          # Get rid of the large file

        end_time = time.time()
        file_time = end_time - start_time

        print(f"{bcolors.OKCYAN}Processing time: {file_time:.2f} seconds{bcolors.ENDC}")
        print("")
        print(f"{'-' * 25}{bcolors.UNDERLINE}Saved: {output_filename}{bcolors.ENDC}{'-' * 25}")
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
        if 'image_pdf' in locals():
            image_pdf.close()
        
        totalcount += 1
        continue

total_time_end = time.time()
elapsed_total_time = total_time_end - total_time_start

if successcount == totalcount:
    print(f"{bcolors.OKGREEN}SUCCESSFULLY CONVERTED {successcount} CROPPED PDF(S) OUT OF {totalcount} CROPPED PDF(S) IN {elapsed_total_time:.2f} SECOND(S) [{display_time(elapsed_total_time)}]{bcolors.ENDC}")
    print(f"{bcolors.OKGREEN}Input directory size: {format_bytes(input_directory_size)} | Output directory size: {format_bytes(get_folder_size_os(output_directory))} {bcolors.ENDC}")

elif 0 < successcount < totalcount:
    print(f"{bcolors.WARNING}WARNING: ALL PDFS COULD NOT BE CROPPED{bcolors.ENDC}")
    print(f"{bcolors.WARNING}SUCCESSFULLY CONVERTED {successcount} CROPPED PDF(S) OUT OF {totalcount} CROPPED PDF(S) IN {elapsed_total_time:.2f} SECOND(S) [{display_time(elapsed_total_time)}]{bcolors.ENDC}")
    print(f"{bcolors.WARNING}Input directory size: {format_bytes(input_directory_size)} | Output directory size: {format_bytes(get_folder_size_os(output_directory))} {bcolors.ENDC}")

else:
    print(f"{bcolors.FAIL}UNABLE TO CACHE ANY PDF(S) from the DIRECTORY{bcolors.ENDC}")