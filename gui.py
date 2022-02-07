import os
from queue import Queue
import sys
import threading

from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QGridLayout, QHBoxLayout, QWidget, QPushButton, QLabel, QLineEdit, QSpinBox, QProgressBar
from PyQt5.QtCore import Qt, QTimer

from helper import *
import main
from settings import *


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.settings = Settings()
        detected_template = detect_file("*.docx", "template")
        if detected_template:
            self.settings.filename_template = detected_template
        detected_spreadsheet = detect_file("*.xlsx")
        if detected_spreadsheet:
            self.settings.filename_spreadsheet = detected_spreadsheet

        self.queue = Queue()

        central_widget = QWidget()
        layout = QGridLayout(central_widget)
        self.setCentralWidget(central_widget)

        label_template = QLabel("Template File:")
        text_template = QLineEdit(self.settings.filename_template)
        text_template.textEdited.connect(self.on_text_template_changed)
        text_template.setAlignment(Qt.AlignRight)
        layout.addWidget(label_template, 0, 0)
        layout.addWidget(text_template, 0, 1)

        label_spreadsheet = QLabel("Spreadsheet File:")
        text_spreadsheet = QLineEdit(self.settings.filename_spreadsheet)
        text_spreadsheet.textEdited.connect(self.on_text_spreadsheet_changed)
        text_spreadsheet.setAlignment(Qt.AlignRight)
        layout.addWidget(label_spreadsheet, 1, 0)
        layout.addWidget(text_spreadsheet, 1, 1)

        label_final = QLabel("Output File:")
        text_final = QLineEdit(self.settings.filename_final)
        text_final.textEdited.connect(self.on_text_final_changed)
        text_final.setAlignment(Qt.AlignRight)
        layout.addWidget(label_final, 2, 0)
        layout.addWidget(text_final, 2, 1)

        label_row = QLabel("Row to Use:")
        text_row = QSpinBox()
        text_row.setMinimum(2)
        text_row.setValue(self.settings.spreadsheet_row+2)
        text_row.valueChanged.connect(self.on_text_row_changed)
        text_row.setAlignment(Qt.AlignRight)
        layout.addWidget(label_row, 3, 0)
        layout.addWidget(text_row, 3, 1)

        label_nested_tables = QLabel("Nested Tables:")
        text_nested_tables = QSpinBox()
        text_nested_tables.setMinimum(0)
        text_nested_tables.setMaximum(100)
        text_nested_tables.setValue(self.settings.nested_tables)
        text_nested_tables.valueChanged.connect(self.on_text_nested_tables_changed)
        text_nested_tables.setAlignment(Qt.AlignRight)
        layout.addWidget(label_nested_tables, 4, 0)
        layout.addWidget(text_nested_tables, 4, 1)

        self.button_run = QPushButton("Run")
        self.button_run.clicked.connect(self.on_button_run_click)
        layout.addWidget(self.button_run, 5, 0, 1, 2)

        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar, 6, 0, 1, 2)

        self.label_status = QLabel()
        self.label_status.setAlignment(Qt.AlignHCenter)
        layout.addWidget(self.label_status, 7, 0, 1, 2)
        
        self.timer = QTimer()
        self.timer.timeout.connect(self.check_queue)
        self.timer.setInterval(10)
    
    def on_text_template_changed(self):
        self.settings.filename_template = self.sender().text()
    
    def on_text_spreadsheet_changed(self):
        self.settings.filename_spreadsheet = self.sender().text()
    
    def on_text_row_changed(self):
        self.settings.spreadsheet_row = self.sender().value() - 2
    
    def on_text_nested_tables_changed(self):
        self.settings.nested_tables = self.sender().value()
    
    def on_text_final_changed(self):
        self.settings.filename_final = self.sender().text()
    
    def on_button_run_click(self):
        if self.settings.filename_final == self.settings.filename_template:
            self.label_status.setText(f"Output file name {self.settings.filename_final} will overwrite the template file.")
            return
        
        self.button_run.setEnabled(False)
        self.thread = threading.Thread(
            target=main.main,
            args=[self.settings, self.queue],
        )
        self.thread.start()
        self.timer.start()
    
    def check_queue(self):
        while not self.queue.empty():
            message = self.queue.get()
            if isinstance(message, str):
                self.label_status.setText(message)
            else:
                self.progress_bar.setValue(message)
        # Thread has stopped.
        if not self.thread.is_alive():
            self.button_run.setEnabled(True)
            self.progress_bar.reset()


if __name__ == "__main__":
    application = QApplication(sys.argv)
    window = MainWindow()
    window.setWindowTitle(PROGRAM_NAME)
    window.show()
    sys.exit(application.exec_())