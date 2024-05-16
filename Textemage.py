import sys
import subprocess
import pandas as pd

# Function to install a Python module using pip
def install_module(module_name):
    subprocess.check_call([sys.executable, "-m", "pip", "install", module_name])

# Function to install pytesseract
def install_pytesseract():
    install_module("pytesseract")

# Function to install openpyxl
def install_openpyxl():
    install_module("openpyxl")

# Function to install pytesseract
def install_pandas():
    install_module("pandas")

# Function to extract text from image using Tesseract OCR
def extract_text_from_image(image_path):
    import pytesseract
    from PIL import Image
    try:
        text = pytesseract.image_to_string(Image.open(image_path))
        print("Extracted text:", text)
        return text.strip()
    except Exception as e:
        print("Error occurred while extracting text from image:", e)
        return None

# Function to parse extracted text and organize it into Excel format
def parse_text_to_excel(text, excel_file, folder_path):
    if text is None:
        print("No text extracted. Excel file will not be created.")
        return

    import openpyxl
    try:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "WhatsApp Status"
        sheet.append(["Contact Name", "Time Seen"])

        # Split text into lines and extract contact name and time seen
        lines = text.split('\n')
        line  = lines[5:]
        line = [item for item in line if item != '']
        Data = {"Contact Name":[],"Time Seen":[]}
        for i in range(0,len(line)):
            if i%2==0:
                temp = Data["Contact Name"]
                temp.append(line[i])
                Data["Contact Name"] = temp
            else:
                temp = Data["Time Seen"]
                temp.append(line[i])
                Data["Time Seen"] = temp

        import os

        df = pd.DataFrame(Data)

        # Check if the folder exists, if not create it
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        # Construct the complete file path
        file_path = os.path.join(folder_path, excel_file)

        # Save DataFrame to Excel using the complete file path
        df.to_excel(file_path, index=False)

        print(f"Data saved to {file_path} successfully!")
    except Exception as e:
        print("Error occurred while saving data to Excel:", e)


# Prompt user to input path to the image file
def get_image_path():
    image_path = input("Enter the path to the image file: ").strip('"')
    folder_path = input("Enter the folder path to save excel file: ").strip('"')
    # folder_path = r"C:\Users\ADMIN\Downloads\test"

    return (image_path,folder_path)

# Main function
def main():
    # Install pytesseract if not already installed
    install_pytesseract()

    # Install openpyxl if not already installed
    install_openpyxl()

    install_pandas()

    # Get path to the image file
    image_path,folder_path = get_image_path()

    # Extract text from image
    extracted_text = extract_text_from_image(image_path)
    print(extracted_text)
    # Path to the output Excel file
    # folder_path = r"C:\Users\ADMIN\Downloads\test"
    excel_file = 'status.xlsx'

    # Parse extracted text and save to Excel
    parse_text_to_excel(extracted_text, excel_file , folder_path)

if __name__ == "__main__":
    main()
