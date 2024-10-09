# Comment Extractor

This is a Python script that extracts comments and their replies from a Microsoft Word document and exports the data to an Excel file.

## Features

- Extracts comments and their replies from the currently open Word document.
- Retrieves the full paragraph that the comment is referencing.
- Exports the extracted data to an Excel file.
- Automatically opens the generated Excel file.
- Packaged as a standalone executable file, making it easy to distribute and use.

## Prerequisites

- Python 3.x
- The following Python libraries:
  - `win32com.client`
  - `pandas`
  - `xlsxwriter`
  - `subprocess`
  - `PyInstaller`

## Usage

1. Open the Word document you want to extract comments from.
2. Run the `comment_extractor.py` script.
3. The script will automatically generate an Excel file named `comment_data.xlsx` in the same directory as the script.
4. The Excel file will be opened automatically after it's created.

## Packaging as an Executable

To create a standalone executable file that can be distributed to other users, follow these steps:

1. Install PyInstaller using pip:

pip install pyinstaller

2. Save the `comment_extractor.py` script.
3. Open a terminal or command prompt and navigate to the directory containing the script.
4. Run the following command to create the executable:

pyinstaller comment_extractor.py

5. The executable file will be created in the `dist` folder within the same directory as the script.
