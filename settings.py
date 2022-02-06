from dataclasses import dataclass

@dataclass
class Settings:
    filename_template : str = "template.docx"
    filename_spreadsheet : str = "data.xlsx"
    filename_final : str = "report.docx"
    
    # The row in the spreadsheet to use. 0 is the row immediately below the header row.
    spreadsheet_row : int = 0

    # Characters that surround placeholders.
    placeholder_prefix : str = "«"
    placeholder_suffix : str = "»"

    # Number of tables within tables to search for.
    nested_tables : int = 1

    ignore_paragraphs : bool = False
    ignore_tables : bool = False
    ignore_headers_footers : bool = False