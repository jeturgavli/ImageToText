import os
import subprocess
import sys
import pandas as pd
from PIL import Image
import pytesseract
import openpyxl
import logging

# This helps in determining if we need to install a module
def is_module_installed(module_name):
    try:
        __import__(module_name)
        return True
    except ImportError:
        return False

# This function is called if a required module is not installed
def install_module(module_name):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", module_name])
    except subprocess.CalledProcessError as e:
        logging.error(f"Error installing {module_name}: {e}")

# Ensures all necessary modules are available for the script to run
def install_required_modules():
    required_modules = ["pytesseract", "openpyxl", "pandas"]
    for module in required_modules:
        if not is_module_installed(module):
            install_module(module)

# Function to extract text from an image 
def extract_text_from_image(img_path):
    try:
        text = pytesseract.image_to_string(Image.open(img_path))
        logging.info(f"Text extraction successful from {img_path}")
        return text.strip()
    except Exception as e:
        logging.error(f"Error occurred while extracting text from {img_path}: {e}")
        return None
'''
This function processes the text to extract 
relevant information and saves it to an Excel file 
'''
def parse_text_to_excel(text, output, imagefiles):
    if text is None:
        logging.warning("No text extracted. Excel file will not be created.")
        return

    try:
        lines = text.split('\n')
        lines = [line for line in lines if line.strip()]  

        #  (at least 6 lines)
        if len(lines) < 6:
            logging.warning(f"Extracted text from {imagefiles} does not contain enough data to process.")
            return

        # Data structure to hold contact names and times seen
        data = {"Contact Name": [], "Time Seen": []}
        for i in range(5, len(lines), 2):
            if i + 1 < len(lines):  
                data["Contact Name"].append(lines[i])
                data["Time Seen"].append(lines[i + 1])
            else:
                logging.warning(f"Incomplete data for contact name {lines[i]} in {imagefiles}")

        if not data["Contact Name"] or not data["Time Seen"]:
            logging.warning(f"No valid data extracted from {imagefiles}. Skipping file creation.")
            return

        df = pd.DataFrame(data)

        excel_file = os.path.join(output, f"{imagefiles.split('.')[0]}.xlsx")

        # Save the DataFrame to an Excel file
        if not os.path.exists(excel_file):
            df.to_excel(excel_file, index=False)
            logging.info(f"Excel file created at {excel_file}")
        else:
            with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
                df.to_excel(writer, index=False, header=False)
            logging.info(f"Data appended to existing Excel file at {excel_file}")
    except PermissionError:
        logging.error(f"Permission denied: '{excel_file}'")
    except Exception as e:
        logging.error(f"Error occurred while saving data to Excel: {e}")

# Main function to drive the script
def main():
    logging.basicConfig(level=logging.INFO)  

    # install modules if not already exist
    install_required_modules()

    # input and output directories
    inputfolder = "images"
    output = "output"

    if not os.path.exists(output):
        os.makedirs(output)

    for imagefiles in os.listdir(inputfolder):
        if imagefiles.lower().endswith((".png", ".jpg", ".jpeg")):
            img_path = os.path.join(inputfolder, imagefiles)

            extracted_text = extract_text_from_image(img_path)

            parse_text_to_excel(extracted_text, output, imagefiles)

if __name__ == "__main__":
    main()
