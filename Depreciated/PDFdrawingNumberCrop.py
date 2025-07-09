import os
import fitz
import time
from pathlib import Path
from DEdependencies import bcolors
from DEdependencies import display_time

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

dirpath = Path(__file__).parent.as_posix()

input_directory = Path(f"{dirpath}/PDFsToProcess")   # For iterating over all the files
output_directory = Path(f"{dirpath}/_workingdata_/_bloatedcache_/")   # Put the finished files here

input_directory.mkdir(parents=True, exist_ok=True)   # Create output directory if it doesn't exist
output_directory.mkdir(parents=True, exist_ok=True)   # Create output directory if it doesn't exist

file_paths = [f for f in input_directory.iterdir() if f.suffix.lower() == '.pdf']   # Get all PDF files in the directory

successcount = 0
totalcount = 0

total_time_start = time.time()
for file_path in file_paths:
    start_time = time.time()     # START TIME OF FILE CROPPING EXECUTION
    try:
        print("")
        dashes = "-" * 25
        print(f"{dashes}{bcolors.UNDERLINE}Processing: {file_path.stem}.pdf{bcolors.ENDC}{dashes}")
        print("")

        doc = fitz.open(str(file_path))
        
        if doc.page_count == 0:
            print(f"{bcolors.WARNING}Warning: {file_path.name} has no pages, skipping...{bcolors.ENDC}")
            doc.close()
            continue

        for page_num in range(doc.page_count):
            page = doc[page_num]

            pagerotation = page.rotation

            page.remove_rotation()

            # Get the MEDIABOX (actual page boundaries)
            page_mediabox = page.mediabox

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
                crop_x0 = (width * 0.5) * 0.88
                crop_y0 = height * 0.92
                crop_x1 = (width * 0.5) * 0.98
                crop_y1 = height * 0.98
            
            else:
                # is rotated, x changes height and y changes length
                crop_x0 = x1 * 0.93
                crop_y0 = y1 * 0.92
                crop_x1 = x1 * 0.99
                crop_y1 = y1 * 0.99

            new_crop_rect = fitz.Rect(crop_x0, crop_y0, crop_x1, crop_y1)   # Crop coordinates for drawing title of each page

            page.set_cropbox(new_crop_rect)

            # page.set_rotation(pagerotation) Not sure why but this isn't required to upright the cropped area, maybe the cropbox auto fixes it;
            # Either way it's not causing issues so I'm not crying
        
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
        # --------------------------- Troubleshooting info ---------------------------

        # Create output filename: original_filename + "DrawlingNumber.pdf"
        original_filename = file_path.stem  # Gets filename without extension
        output_filename = f"{original_filename}-drawingno.pdf"
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

total_time_end = time.time()
elapsed_total_time = total_time_end - total_time_start
print(f"{bcolors.OKGREEN}SUCCESSFULLY EXTRACTED DRAWING NUMBER(S) FROM {successcount} FILE(S) OUT OF {totalcount} FILE(S) IN {elapsed_total_time:.2f} SECOND(S) [{display_time(elapsed_total_time)}]{bcolors.ENDC}")