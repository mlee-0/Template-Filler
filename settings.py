from dataclasses import dataclass

@dataclass
class Settings:
    filename_template : str = "template.docx"
    filename_data : str = "data.xlsx"
    filename_final : str = "report.docx"
    # The row in the spreadsheet to use. 0 is the row immediately below the header row.
    spreadsheet_row : int = 0