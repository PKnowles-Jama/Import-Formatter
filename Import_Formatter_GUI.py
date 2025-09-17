# Import_Formatter_GUI

import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QFormLayout, QRadioButton, QPushButton, QLineEdit, QFileDialog, QLabel, QComboBox, QTextEdit, QHBoxLayout)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon
from Import_Formatter_Functions import KeywordParseExcel, DefaultParseExcel
from PermanentHeader import permanent_header
from NextButton import NextButton
import io
from contextlib import redirect_stdout

class ImportFormatterGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Import Formatter")
        screen = QApplication.primaryScreen()
        screen_geometry = screen.geometry()
        window_width = 800
        window_height = 100
        x = (screen_geometry.width() - window_width) // 2
        y = (screen_geometry.height() - window_height) // 2
        self.setGeometry(x, y, window_width, window_height)

        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, 'jama_logo_icon.png')
        self.setWindowIcon(QIcon(icon_path))

        self.main_app_layout = QVBoxLayout()
        self.setLayout(self.main_app_layout)

        header_layout, separator = permanent_header("Import Formatter", 'jama_logo.png')
        self.main_app_layout.addLayout(header_layout)
        self.main_app_layout.addWidget(separator)

        self.dynamic_content_layout = QVBoxLayout()
        self.main_app_layout.addLayout(self.dynamic_content_layout)
        self.main_app_layout.addStretch()

        self.HomePage()

    def HomePage(self):
        form_layout = QFormLayout()

        self.excel = QRadioButton("Excel")
        # TODO add Word, CSV, etc as other functions are added to the script
        self.excel.setChecked(True)

        radio_button_layout = QHBoxLayout()
        # TODO add Word, CSV, etc as other functions are added to the script
        radio_button_layout.addWidget(self.excel)
        radio_button_layout.addStretch()

        form_layout.addRow("Select Input File Format:", radio_button_layout)

        self.submit_button = NextButton("Submit", True)

        self.dynamic_content_layout.addLayout(form_layout)
        self.dynamic_content_layout.addWidget(self.submit_button)
        self.dynamic_content_layout.addStretch()

        if self.excel.isChecked():
            self.file_format = '.xlsx' 
        else: # TODO update for other options
            self.file_format = '.xlsx' 

        self.submit_button.clicked.connect(self.DetailsPage)

    def DetailsPage(self):
        self.clearLayout(self.dynamic_content_layout)
        layout = QVBoxLayout()

        self.select_file_button = QPushButton("Select File")
        self.select_file_button.clicked.connect(self.select_file)
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.select_file_button)
        self.select_file_button.setStyleSheet("background-color: #0052CC; color: white;") # Set the button color
        self.file_path_label = QLabel("No file selected")
        button_layout.addWidget(self.file_path_label)
        layout.addLayout(button_layout) 

        self.submit_button = NextButton("Submit", True)

        self.dynamic_content_layout.addLayout(layout)
        self.dynamic_content_layout.addWidget(self.submit_button)
        self.dynamic_content_layout.addStretch()

        self.submit_button.clicked.connect(self.FinalDetailsPage)

    def FinalDetailsPage(self):
        self.clearLayout(self.dynamic_content_layout)
        layout = QFormLayout()

        try:
            # First, try to read the file using the default pandas behavior
            # Pandas will try openpyxl for .xlsx and xlrd for .xls
            df = pd.read_excel(self.file_path)
        except Exception as e:
            # If the default behavior fails with a No engine error, it might be due to a specific file type issue.
            if "No engine for filetype: 'xls'" in str(e):
                try:
                    # Explicitly try the xlrd engine for older files
                    df = pd.read_excel(self.file_path, engine='xlrd')
                except Exception as inner_e:
                    # If even the explicit engine fails, raise a more descriptive error.
                    self.handle_error(f"Failed to process .xls file with xlrd. Please ensure the file is not corrupted. Error: {inner_e}")
                    return
            else:
                # For all other errors, use a generic handler.
                self.handle_error(f"An unexpected error occurred: {e}")
                return

        # If the file reading was successful, proceed with the rest of the function
        column_names = df.columns.tolist()

        # Add QComboBox that lists the names of each column
        self.column_combo = QComboBox()
        self.column_combo.addItems(column_names)
        layout.addRow("Select Column:", self.column_combo)

        # Create Parsing Selection Radio Buttons
        self.keyword = QRadioButton("Keyword")
        self.standard = QRadioButton("Default Formats")
        self.standard.setChecked(True)
        radio_container = QWidget()
        radio_layout = QHBoxLayout(radio_container)
        radio_layout.setContentsMargins(0, 0, 0, 0)
        radio_layout.addWidget(self.standard)
        radio_layout.addWidget(self.keyword)
        radio_layout.addStretch() 
        layout.addRow("Select Parsing Method:", radio_container)

        self.keyword_label = QLabel("Enter Keyword:")
        self.word_input = QLineEdit()
        self.word_input.setPlaceholderText("This is NOT case sensitive.")
        layout.addRow(self.keyword_label, self.word_input)
        self.keyword.toggled.connect(self._toggle_keyword_widgets)
        self._toggle_keyword_widgets()
        
        self.submit_button = NextButton("Submit", True)

        self.dynamic_content_layout.addLayout(layout)
        self.dynamic_content_layout.addWidget(self.submit_button)
        self.dynamic_content_layout.addStretch()
        
        # Add the QTextEdit to display the output
        self.output_text_box = QTextEdit()
        self.output_text_box.setReadOnly(True)
        self.dynamic_content_layout.addWidget(self.output_text_box)

        self.submit_button.clicked.connect(self.on_submit)

    def _toggle_keyword_widgets(self):
        """
        Shows or hides the keyword input widgets based on which
        radio button is selected.
        """
        is_visible = self.keyword.isChecked()
        self.keyword_label.setVisible(is_visible)
        self.word_input.setVisible(is_visible)

    def handle_error(self, message):
        """
        A helper function to clear the layout and display an error message.
        """
        self.clearLayout(self.dynamic_content_layout)
        error_label = QLabel(message)
        error_label.setWordWrap(True)
        self.dynamic_content_layout.addWidget(error_label)
        self.dynamic_content_layout.addStretch()

    def on_submit(self):
        """
        This function is a simple wrapper that gets the selected values from the UI
        and passes them to the ParseExcel function.
        """
        # Create a new string stream to capture stdout
        sys.stdout = stream = io.StringIO()
        
        try:
            # Get the selected column name from the QComboBox
            selected_column = self.column_combo.currentText()
            
            # Get the keyword from the QLineEdit
            keyword = self.word_input.text()

            if self.keyword.isChecked():
                # Call the ParseExcel function with the collected values
                KeywordParseExcel(self.file_path, keyword, selected_column)
            else:
                DefaultParseExcel(self.file_path, selected_column)

        finally:
            # Restore stdout to its original state
            sys.stdout = sys.__stdout__

        # Get the captured output from the stream
        output = stream.getvalue()
        
        # Set the text of the QTextEdit to the captured output
        self.output_text_box.setText(output)
        
    def clearLayout(self, layout):
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
                else:
                    self.clearLayout(item.layout())

    def select_file(self):
        #
        # This function does the following:
        #   Open file explorer.
        #   Collect user's file name & path.
        #   Only allow user to select the previously determined file type
        #
        file_dialog = QFileDialog()
        if self.file_format == '.xlsx':
            file_filter = "Excel Files (*.xlsx *.xls)"
        else:
            file_filter = f"Files (*{self.file_format})"
        file_path, _ = file_dialog.getOpenFileName(self, "Select File", "", file_filter)
        if file_path:
            self.file_path = file_path
            self.file_path_label.setText(f"Selected: {self.file_path}")
            self.select_file_button.setStyleSheet("background-color: #53575A; color: white;") # Set the button color

if __name__ == "__main__":
    # Run the app when running this file
    app = QApplication(sys.argv)
    window = ImportFormatterGUI()
    window.show()
    sys.exit(app.exec())