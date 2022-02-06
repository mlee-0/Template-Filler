# Template Completer
A GUI program that fills in merge fields in a Microsoft Word document using data in a Microsoft Excel spreadsheet.

## Usage
1. Create an Excel spreadsheet to be linked to the Word document.
    - Open Word
    - Click the Mailings tab
    - "Select Recipients"
    - "Use an Existing List..."
    - Find and open the Excel spreadsheet
2. Insert merge fields in the Word document.
    - Place the cursor where a merge field should be inserted
    - Click the Mailings tab
    - "Insert Merge Field"
        - The fields that appear are the header labels from the Excel spreadsheet.
    - Capitalize any merge fields, if needed
        - Right click a merge field
        - "Edit Field..."
        - Under "Format", click "First capital"

## Features
* ~~Automatically omits the *s* in instances of *'s* if preceded by a name that ends in *s*.~~

## Requirements
Written in Python 3.9 on Windows.

Enter `pip install requirements.txt` in the console to automatically download dependencies.
* python-docx
* pandas
* openpyxl
* PyQt5