import os
import re

from docx import Document
import pandas as pd

from student import Student


FILENAME_TEMPLATE = "template.docx"
FILENAME_FINAL = "report.docx"
FILENAME_DATA = "data.xlsx"

# The row in the spreadsheet to use. 0 is the row immediately below the header row.
SPREADSHEET_ROW = 0

PLACEHOLDER_PREFIX = "«"
PLACEHOLDER_SUFFIX = "»"
# List of apostrophe characters to search for. The first character in this list will be used in text replacements.
APOSTROPHES = ["’", "'"]

# Highlight filled in text for debugging.
highlight_filled_text = True


if __name__ == "__main__":
    # Open the template.
    document = Document(FILENAME_TEMPLATE)

    # Open the spreadsheet.
    df = pd.read_excel(
        FILENAME_DATA,
        sheet_name=0,  # Get first sheet only
        header=0,  # Get labels from first row
    )
    df = df.add_prefix(PLACEHOLDER_PREFIX)
    df = df.add_suffix(PLACEHOLDER_SUFFIX)
    # Replace any spaces in column headers with underscores to match the text in the template.
    df = df.rename(columns=lambda string: string.replace(" ", "_"))
    
    # Store information in a Student object.
    student = Student(
        name_first=df.at[SPREADSHEET_ROW, df.columns[0]],
        name_middle=df.at[SPREADSHEET_ROW, df.columns[1]],
        name_last=df.at[SPREADSHEET_ROW, df.columns[2]],
        gender=df.at[SPREADSHEET_ROW, df.columns[4]],
        grade=df.at[SPREADSHEET_ROW, df.columns[5]],
        age=df.at[SPREADSHEET_ROW, df.columns[6]],
        birthday=df.at[SPREADSHEET_ROW, df.columns[7]],
    )

    # Store the data from the spreadsheet in a dictionary.
    data = {label: df.at[SPREADSHEET_ROW, label] for label in df.columns}
    # Search for instances of 's following the student's name and omit the "s" if the first name ends with "s".
    if student.name_first[-1].lower() == "s":
        # Insert this item at the beginning so that instances of 's are searched before searching for instances of first name without 's.
        data = {
            **{f"{df.columns[0]}({'|'.join(APOSTROPHES)})s": f"{student.name_first}{APOSTROPHES[0]}"},
            **data,
        }

    # Iterate over the paragraphs in the template.
    print("Reading paragraphs...")
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            for placeholder, replacement in data.items():
                run.text = re.sub(
                    placeholder,
                    str(replacement),
                    run.text,
                )

    # Iterate over the tables in the template.
    print("Reading tables...")
    for table in document.tables:
        for column in table.columns:
            for cell in column.cells:
                if PLACEHOLDER_PREFIX in cell.text or PLACEHOLDER_SUFFIX in cell.text:
                    for placeholder, replacement in data.items():
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                cell.text = re.sub(
                                    placeholder,
                                    str(replacement),
                                    cell.text,
                                )
    
    # Save the modified template as a new file.
    document.save(FILENAME_FINAL)