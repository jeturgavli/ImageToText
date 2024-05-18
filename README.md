# Text Extraction and Excel Parsing from Images

This Python script extracts text from images using Tesseract OCR and organizes it into an Excel file.

## Features

- **Automated Installation:** Checks for required Python modules (`pytesseract`, `openpyxl`, `pandas`) and installs them if missing.
- **Text Extraction:** Utilizes Tesseract OCR to extract text from images.
- **Data Parsing:** Parses extracted text to extract contact names and times seen, organizing them into an Excel file.
- **Logging:** Logs informative messages, warnings, and errors for better tracking and debugging.
- **User Interaction:** Prompts the user for image and output folder paths, allowing for interactive usage.

## Usage

1. Ensure Python is installed.
2. Install Tesseract OCR:
   - **Windows:**
     - Download the installer from [https://github.com/UB-Mannheim/tesseract/wiki](https://github.com/UB-Mannheim/tesseract/wiki).
     - Run the installer and follow the installation instructions.
     - Add the Tesseract installation directory to the system's PATH environment variable.
     - [click here to watch how install Tesseract Ocr for windows](https://www.youtube.com/watch?v=2kWvk4C1pMo)
   - **Linux:**
     - Use your package manager to install Tesseract OCR. For example, on Ubuntu:
       ```
       sudo apt-get update
       sudo apt-get install tesseract-ocr
       ```
   - **macOS:**
     - Install Tesseract OCR using Homebrew:
       ```
       brew install tesseract
       ```
3. Clone or download the repository.
4. Place images to be processed in the `images` folder.
5. Run the script (`main.py`).
6. Follow the prompts to input image and output folder paths.
7. View the generated Excel files in the `output` folder.

## Dependencies

- Python 3.x
- Tesseract OCR
- Required Python modules: `pytesseract`, `openpyxl`, `pandas`

## Author
[LAKSHMI](https://github.com/LAKSHMI-DEVI-REDDY-MALLI)

## Contribution
[JETUR GAVLI](https://github.com/jeturgavli)

## License

This project is licensed under the [MIT License](LICENSE).
