from dataclasses import dataclass


PROGRAM_NAME = "Template Filler"
VERSION_MAJOR = 1
VERSION_MINOR = 0
VERSION_PATCH = 0

# Characters that surround placeholders.
PLACEHOLDER_PREFIX : str = "«"
PLACEHOLDER_SUFFIX : str = "»"

@dataclass
class Settings:
    filename_template : str = "template.docx"
    filename_spreadsheet : str = "data.xlsx"
    filename_final : str = "report.docx"
    
    # The row in the spreadsheet to use. 0 is the row immediately below the header row.
    spreadsheet_row : int = 0

    # Number of tables within tables to search for.
    nested_tables : int = 1

    ignore_paragraphs : bool = False
    ignore_tables : bool = False
    ignore_headers_footers : bool = False

    date_format = r"%m/%d/%Y"