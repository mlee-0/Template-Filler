import os
import re

from docx import Document
import pandas as pd

from student import Student
from settings import Settings


def main(settings: Settings):
    PLACEHOLDER_PREFIX = "«"
    PLACEHOLDER_SUFFIX = "»"
    # List of apostrophe characters to search for. The first character in this list will be used in text replacements.
    APOSTROPHES = ["’", "'"]

    # Open the template.
    document = Document(settings.filename_template)

    # Open the spreadsheet.
    df = pd.read_excel(
        settings.filename_data,
        sheet_name=0,  # Get first sheet only
        header=0,  # Get labels from first row
    )
    df = df.add_prefix(PLACEHOLDER_PREFIX)
    df = df.add_suffix(PLACEHOLDER_SUFFIX)
    # Replace any spaces in column headers with underscores to match the text in the template.
    df = df.rename(columns=lambda string: string.replace(" ", "_"))
    
    # Store information in a Student object.
    student = Student(
        name_first=df.at[settings.spreadsheet_row, df.columns[0]],
        name_middle=df.at[settings.spreadsheet_row, df.columns[1]],
        name_last=df.at[settings.spreadsheet_row, df.columns[2]],
        gender=df.at[settings.spreadsheet_row, df.columns[4]],
        grade=df.at[settings.spreadsheet_row, df.columns[5]],
        age=df.at[settings.spreadsheet_row, df.columns[6]],
        birthday=df.at[settings.spreadsheet_row, df.columns[7]],
    )

    # Store the data from the spreadsheet in a dictionary.
    data = {label: df.at[settings.spreadsheet_row, label] for label in df.columns}
    # Insert pronouns.
    data[f"{PLACEHOLDER_PREFIX}he/she{PLACEHOLDER_SUFFIX}"] = student.pronoun_subject
    data[f"{PLACEHOLDER_PREFIX}him/her{PLACEHOLDER_SUFFIX}"] = student.pronoun_object
    data[f"{PLACEHOLDER_PREFIX}his/her{PLACEHOLDER_SUFFIX}"] = student.pronoun_possessive_dependent
    data[f"{PLACEHOLDER_PREFIX}his/hers{PLACEHOLDER_SUFFIX}"] = student.pronoun_possessive_independent
    data[f"{PLACEHOLDER_PREFIX}himself/herself{PLACEHOLDER_SUFFIX}"] = student.pronoun_reflexive

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
    document.save(settings.filename_final)
    print(f"Wrote {settings.filename_final}")

if __name__ == "__main__":
    settings = Settings()
    main(settings)