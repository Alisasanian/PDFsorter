import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import shutil
import fitz
import time
import io
import re
import tempfile
from pathlib import Path
from PIL import Image
from contextlib import redirect_stderr
from DEdependencies import bcolors
from DEdependencies import display_time
from DEdependencies import printProgressBar

"""     #1
We want to crop out a way to index and regex out the useful sheets to reduce OCR load.
By efforlessly (relatively) cropping the PDF to only showcase pixels representing the drawing number, we quickly OCR the titles

This script is VERY FAST at cropping out the drawing numbers, however this is at the cost of taking up more cached space; essentially 1:1 or greater.
This is because the "cropped" pdfs that are saved do not actually store less data than their uncropped counterparts, thus filling up space quickly.
Overall, there shouldn't be a problem running the script as long as there is enough space and the cropped files are deleted after being used.
If a more permanet solution is required to store the cropped PDFs for future reference (Future cached storage), here is a potential solution:
Create a child script that runs tangentially or sequentially while using this script. The other script will "screenshot" a picture of each pdf page
and store the picutres as pages for a new pdf, and after completion the "cropped" pdf is deleted. This is more space efficient at the cost of performance.
*Note that the screenshot method is not perfect as you will pruge XREF data and other non-visible data from the PDF. I am not liable for any data lost because of my idea, use this at your own risk!

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

def flattenPDF(pdf_path, start_time):
    doc = fitz.open(str(pdf_path))     # Open the broken PDF

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

    bwoken_file_name = f"{pdf_path.stem}-bwoken.pdf"
    image_pdf.save(str(bwoken_directory / bwoken_file_name))        # Save the new PDF with images
    image_pdf.close()
    doc.close()

    end_time = time.time()
    file_time = end_time - start_time

    print("")
    print(f"{bcolors.OKGREEN}SUCCESSFULLY REVIVED THE BROKEN PDF")
    print(f"{bcolors.OKCYAN}Processing time: {file_time:.2f} seconds{bcolors.ENDC}")
    print("")
    print(f"{'-' * 25}{bcolors.UNDERLINE}Saved: {bwoken_file_name}{bcolors.ENDC}{'-' * 25}")
    print("")

# -------------------------------- END --------------------------------

dirpath = Path(__file__).parent.as_posix()

input_directory = Path(f"{dirpath}/PDFsToProcess")   # For iterating over all the files
output_directory = Path(f"{dirpath}/_workingdata_/_bloatedcache_/")   # Put the finished files here
bwoken_directory = Path(f"{dirpath}/_workingdata_/_bwokenPDFs_/")        # :(

input_directory.mkdir(parents=True, exist_ok=True)   # Create input directory if it doesn't exist
output_directory.mkdir(parents=True, exist_ok=True)   # Create output directory if it doesn't exist'
bwoken_directory.mkdir(parents=True, exist_ok=True)    # :(

file_paths = [f for f in input_directory.iterdir() if f.suffix.lower() == '.pdf']   # Get all PDF files in the directory

successcount = 0
totalcount = 0

error_pattern = re.compile(r'(cannot|rect|code|MuPDF error:|format error:)')

total_time_start = time.time()
def cropPDF(file_paths):
    global successcount, totalcount

    for file_path in file_paths:
        start_time = time.time()     # START TIME OF FILE CROPPING EXECUTION
        try:
            fitz.TOOLS.reset_mupdf_warnings()  # Clear any previous warnings/errors
            warning_list = []
            print("")
            dashes = "-" * 25
            print(f"{dashes}{bcolors.UNDERLINE}Processing: {file_path.stem}.pdf{bcolors.ENDC}{dashes}")
            print("")

            doc = fitz.open(str(file_path))

            if doc.page_count == 0:
                print(f"{bcolors.WARNING}Warning: {file_path.name} has no pages, skipping...{bcolors.ENDC}")
                doc.close()
                continue

            was_flattened = False  # Track if PDF was flattened

            for page_num in range(doc.page_count):
                page = doc[page_num]

                pagerotation = page.rotation

                page.remove_rotation()

                page_mediabox = page.mediabox       # Get the MEDIABOX (actual page boundaries)

                # mediabox data
                x0 = page_mediabox.x0
                x1 = page_mediabox.x1
                y0 = page_mediabox.y0
                y1 = page_mediabox.y1
                width = page_mediabox.width
                height = page_mediabox.height
                
                
                if x0 < 0 and y0 < 0:
                    # Calculate crop coordinates based on (0,0) center
                    # we start x from the middle, but height from the top downwards
                    crop_x0 = (width * 0.5) * 0.80
                    crop_y0 = height * 0.90
                    crop_x1 = (width * 0.5) * 0.99
                    crop_y1 = height * 0.99
                
                else:
                    # is rotated, x changes height and y changes length
                    crop_x0 = x1 * 0.85
                    crop_y0 = y1 * 0.85
                    crop_x1 = x1 * 0.99
                    crop_y1 = y1 * 0.99

                new_crop_rect = fitz.Rect(crop_x0, crop_y0, crop_x1, crop_y1)   # Crop coordinates for drawing title of each page

                mupdf_warns1 = fitz.TOOLS.mupdf_warnings()     # Check for warnings or errors generated during open
                if mupdf_warns1:
                    warning_list.append(mupdf_warns1)

                page.set_cropbox(new_crop_rect)

                mupdf_warns2 = fitz.TOOLS.mupdf_warnings()     # Check for warnings or errors generated during open
                if mupdf_warns2:
                    warning_list.append(mupdf_warns2)

                # page.set_rotation(pagerotation) Not sure why but this isn't required to upright the cropped area, maybe the cropbox auto fixes it;
                # Either way it's not causing issues so I'm not crying
            
            # print(warning_list)

            for warning in warning_list:
                warning_instance = str(warning)

                error_check = error_pattern.search(warning_instance)

                # Only flatten if it's critical and not ignorable
                if error_check:
                    print("")
                    print(f"{bcolors.OKCYAN}| ------------------ HOLD... CRITICAL MuPyPDF ERROR DETECTED... FLATTENING PDF BEFORE PROCEEDING ------------------ |{bcolors.ENDC}")
                    print("")
                    doc.close()
                    flattenPDF(file_path, start_time)
                    successcount += 1
                    totalcount += 1
                    was_flattened = True
                    break

            
            # --------------------------- Troubleshooting info ---------------------------
            '''
            page = doc[0]  # Check first page info

            print(f"Page rotation: {pagerotation}")

            print(f"x0 = {x0}")
            print(f"x1 = {x1}")
            print(f"y0 = {y0}")
            print(f"y1 = {y1}")
            print(f"Width: {width}")
            print(f"Height: {height}")

            print(f"Coords: {crop_x0}, {crop_y0}, {crop_x1}, {crop_y1}")
            '''
            # ----------------------------------- END ------------------------------------

            if was_flattened == False:
                # Create output filename: original_filename + "DrawlingNumber.pdf"
                original_filename = file_path.stem  # Gets filename without extension
                if re.search(r'-bwoken$', original_filename):
                    output_filename = re.sub(r'-bwoken$', '-drawingno', original_filename) + ".pdf"     # Remove "-bwoken" suffix and replace with "-drawingno"
                else:
                    output_filename = f"{original_filename}-drawingno.pdf"      # Just add "-drawingno" as before
                
                output_path = output_directory / output_filename

                # Save the cropped PDF
                doc.save(str(output_path))
                doc.close()

                end_time = time.time()
                file_time = end_time - start_time   # Time taken to crop this file

                print(f"{bcolors.OKCYAN}Processing time: {file_time:.2f} seconds{bcolors.ENDC}")
                
                print("")
                print(f"{dashes}{bcolors.UNDERLINE}Saved: {output_filename}{bcolors.ENDC}{dashes}")
                print("")

                # Pray this worked

                successcount += 1
                totalcount += 1

        except Exception as e:
            print("")
            print("-" * 75)
            print(f"{bcolors.FAIL}Error processing {file_path.name}: {str(e)}{bcolors.ENDC}")

            print(f"Cropbox position: {page.cropbox_position}")
            print(f"Mediabox position: {page_mediabox}")
            print(f"crop_x0 = {crop_x0}")
            print(f"crop_x1 = {crop_x1}")
            print(f"crop_y0 = {crop_y0}")
            print(f"crop_y1 = {crop_y1}")
            print("")
            print(f"Page rotation: {pagerotation}")

            print("-" * 75)
            print("")
            if 'doc' in locals():
                doc.close()
            
            totalcount += 1
            continue

cropPDF(file_paths)

bwoken_paths = [f for f in bwoken_directory.iterdir()]   # Get all PDF files in the directory whilst updating the freshly added bwoken PDFs

cropPDF(bwoken_paths)

try:
    shutil.rmtree(bwoken_directory)
    print(f"{bcolors.OKGREEN}Successfully cleaned up temporary directory: {bwoken_directory}{bcolors.ENDC}")
except PermissionError as e:
    print(f"{bcolors.WARNING}Warning: Could not remove temporary directory {bwoken_directory}: {e}{bcolors.ENDC}")
except Exception as e:
    print(f"{bcolors.FAIL}Error removing temporary directory: {e}{bcolors.ENDC}")

total_time_end = time.time()
elapsed_total_time = total_time_end - total_time_start
print(f"{bcolors.OKGREEN}SUCCESSFULLY EXTRACTED DRAWING NUMBER(S) FROM {successcount} FILE(S) OUT OF {totalcount} FILE(S) IN {elapsed_total_time:.2f} SECOND(S) [{display_time(elapsed_total_time)}]{bcolors.ENDC}")