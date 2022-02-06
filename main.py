import os
from queue import Queue
import re

from docx import Document
import pandas as pd

from helper import *
from student import Student
from settings import Settings


MESSAGE_START_PARAGRAPHS = "Reading paragraphs..."
MESSAGE_START_TABLES = "Reading tables..."
MESSAGE_START_HEADERS_FOOTERS = "Reading headers and footers..."

def replace_in_paragraph(paragraph, dictionary, settings):
    if is_placeholder_in_text(paragraph.text, settings.placeholder_prefix, settings.placeholder_suffix):
        for run in paragraph.runs:
            for placeholder, replacement in dictionary.items():
                run.text = replace_text(placeholder, replacement, run.text)
    return paragraph

def replace_in_table(table, dictionary, settings, recursive_calls=1):
    for column in table.columns:
        for cell in column.cells:
            if cell.tables and recursive_calls > 0:
                for nested_table in cell.tables:
                    nested_table = replace_in_table(nested_table, dictionary, settings, recursive_calls-1)
            else:
                for paragraph in cell.paragraphs:
                    paragraph = replace_in_paragraph(paragraph, dictionary, settings)
    return table

def main(settings: Settings, queue : Queue = None):
    # List of apostrophe characters to search for. The first character in this list will be used in text replacements.
    APOSTROPHES = ["’", "'"]

    # Open the template.
    document = Document(settings.filename_template)

    # Open the spreadsheet.
    df = pd.read_excel(
        settings.filename_spreadsheet,
        sheet_name=0,  # Get first sheet only
        header=0,  # Get labels from first row
    )
    df = df.add_prefix(settings.placeholder_prefix)
    df = df.add_suffix(settings.placeholder_suffix)
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
    data = {label: str(df.at[settings.spreadsheet_row, label]) for label in df.columns}
    # Insert pronouns.
    data[f"{settings.placeholder_prefix}he_she{settings.placeholder_suffix}"] = student.pronoun_subject
    data[f"{settings.placeholder_prefix}him_her{settings.placeholder_suffix}"] = student.pronoun_object
    data[f"{settings.placeholder_prefix}his_her{settings.placeholder_suffix}"] = student.pronoun_possessive_dependent
    data[f"{settings.placeholder_prefix}his_hers{settings.placeholder_suffix}"] = student.pronoun_possessive_independent
    data[f"{settings.placeholder_prefix}himself_herself{settings.placeholder_suffix}"] = student.pronoun_reflexive
    # Check that no header label contains upper case letters.
    assert are_keys_lowercase(data), f"Make sure all headers in {settings.filename_spreadsheet} are lower case and that there are no empty columns."

    # Search for instances of 's following the student's name and omit the "s" if the first name ends with "s".
    if student.name_first[-1].lower() == "s":
        # Insert this item at the beginning so that instances of 's are searched before searching for instances of first name without 's.
        data = {
            **{f"{df.columns[0]}({'|'.join(APOSTROPHES)})s": f"{student.name_first}{APOSTROPHES[0]}"},
            **data,
        }

    # Iterate over the paragraphs in the template.
    print(MESSAGE_START_PARAGRAPHS)
    if queue:
        queue.put(MESSAGE_START_PARAGRAPHS)
    paragraph_count = len(document.paragraphs)
    for i, paragraph in enumerate(document.paragraphs, 1):
        paragraph = replace_in_paragraph(paragraph, data, settings)
        if queue:
            queue.put(round(100 * i / paragraph_count))

    # Iterate over the tables in the template.
    print(MESSAGE_START_TABLES)
    if queue:
        queue.put(MESSAGE_START_TABLES)
    table_count = len(document.tables)
    for i, table in enumerate(document.tables, 1):
        table = replace_in_table(table, data, settings, settings.nested_tables)
        if queue:
            queue.put(round(100 * i / table_count))
    
    # Save the modified template as a new file.
    document.save(settings.filename_final)
    message_done = f"Wrote {settings.filename_final}."
    print(message_done)
    if queue:
        queue.put(message_done)

if __name__ == "__main__":
    settings = Settings()
    main(settings)