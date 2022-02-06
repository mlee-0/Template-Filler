import sys
import threading

from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QGridLayout, QHBoxLayout, QWidget, QPushButton, QLabel, QLineEdit, QSpinBox, QProgressBar

import main
from settings import Settings


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.settings = Settings()

        layout = QGridLayout(self)

        label_template = QLabel("Template File:")
        text_template = QLineEdit("template.docx")
        text_template.textEdited.connect(self.on_text_template_changed)
        layout.addWidget(label_template, 0, 0)
        layout.addWidget(text_template, 0, 1)

        label_spreadsheet = QLabel("Spreadsheet File:")
        text_spreadsheet = QLineEdit("data.xlsx")
        text_spreadsheet.textEdited.connect(self.on_text_spreadsheet_changed)
        layout.addWidget(label_spreadsheet, 1, 0)
        layout.addWidget(text_spreadsheet, 1, 1)

        label_row = QLabel("Row to Use:")
        text_row = QSpinBox()
        text_row.setMinimum(0)
        text_row.valueChanged.connect(self.on_text_row_changed)
        layout.addWidget(label_row, 2, 0)
        layout.addWidget(text_row, 2, 1)

        label_final = QLabel("Output File:")
        text_final = QLineEdit("report.docx")
        text_final.textEdited.connect(self.on_text_final_changed)
        layout.addWidget(label_final, 3, 0)
        layout.addWidget(text_final, 3, 1)

        button_run = QPushButton("Run")
        button_run.clicked.connect(self.on_button_run_click)
        layout.addWidget(button_run, 4, 0, 1, 2)

        progress_bar = QProgressBar()
        layout.addWidget(progress_bar, 5, 0, 1, 2)
        
        # self.thread = threading.Thread(
        #     target=main.main,
        #     args=[self.settings]
        # )
    
    def on_text_template_changed(self):
        self.settings.filename_template = self.sender().text()
    
    def on_text_spreadsheet_changed(self):
        self.settings.filename_spreadsheet = self.sender().text()
    
    def on_text_row_changed(self):
        self.settings.spreadsheet_row = int(self.sender().text())
    
    def on_text_final_changed(self):
        self.settings.filename_final = self.sender().text()
    
    def on_button_run_click(self):
        thread = threading.Thread(
            target=main.main,
            args=[self.settings]
        )
        thread.start()
        # self.thread.start()


if __name__ == "__main__":
    application = QApplication(sys.argv)
    window = MainWindow()
    window.setWindowTitle("Template Completer")
    window.show()
    sys.exit(application.exec_())