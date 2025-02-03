# Word2PDF-Seal

![Python](https://img.shields.io/badge/Python-3.x-blue)
![License](https://img.shields.io/badge/License-MIT-green)

A portable Python program that converts Word documents (.docx) to PDF and adds a seal (image) to the bottom-right corner of the first page of each PDF.

---

## Features

- **Full automation**: Converts all `.docx` files in a folder to PDF.
- **Custom seal**: Adds an image (seal) to the bottom-right corner of the first page of each PDF.
- **Portability**: No installation or additional configuration required. Simply run the program on any Windows PC.
- **Automatic folder creation**: The input folder (`CONFIRMED GUIDES WITHOUT SEAL`) and output folder (`CONFIRMED GUIDES`) are created automatically if they don't exist.

---

## Requirements

- **Operating system**: Windows 10 or higher.
- **Dependencies**:
  - Python 3.x (only required if you want to modify or test the source code).
  - Python libraries: `python-docx`, `pypandoc`, `reportlab`, `Pillow`, `PyPDF2`, `docx2pdf`.

> **Note**: The generated executable does not require the user to have Python installed.

---

## Installation

1. Download the executable file from the [releases section](https://github.com/achf16/Word2PDF-Seal/releases).
2. Place the executable and the seal image (`seal.png`) in the same folder.
3. Run the program. The following folders will be created automatically:
   - `CONFIRMED GUIDES WITHOUT SEAL`: Place your `.docx` files here for conversion.
   - `CONFIRMED GUIDES`: Processed PDF files with the seal will be saved here.

---

## Usage

1. Place your `.docx` files in the `CONFIRMED GUIDES WITHOUT SEAL` folder.
2. Run the program.
3. The PDF files with the seal added will appear in the `CONFIRMED GUIDES` folder.

---

## Customization

- **Change the seal image**:
  - Replace the `seal.png` file with your own image. Ensure the file name remains `seal.png`.

- **Modify the code**:
  - If you want to change the position or size of the seal, edit the code in the `add_seal_to_pdf()` function.

---

## Development

If you want to modify or improve the program, follow these steps:

1. Clone this repository:
   ```bash
   git clone https://github.com/achf16/Word2PDF-Seal.git
2. Install the dependencies:
   ```bash
    pip install -r requirements.txt
3. Run the program:
   ```bash
   python your_script.py
5. Generate a new executable (requires Windows):
   ```bash
   pyinstaller --onefile your_script.py

---
##License 

This project is licensed under the MIT License . You are free to use, modify, and distribute it. 

---
##Contributions 

Contributions are welcome! If you find any issues or have ideas to improve the program, open an issue or submit a pull request. 

---
#Author 

Developed by Abel Chávez
Contact: achf16@gmail.com
GitHub: https://github.com/achf16
