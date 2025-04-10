"""
GUI application for the DOCX Table Converter.
"""

import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                            QHBoxLayout, QPushButton, QLabel, QFileDialog,
                            QSpinBox, QComboBox, QTextEdit, QTableWidget,
                            QTableWidgetItem, QCheckBox, QGroupBox, QMessageBox,
                            QDialog, QLineEdit, QFormLayout)
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QFont, QIcon
from .core import write_table_to_docx, read_table_from_file, parse_clipboard_data


class TablePreviewDialog(QDialog):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.df = df
        self.setWindowTitle("Table Preview")
        self.setMinimumSize(800, 600)
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Create table widget
        table = QTableWidget()
        table.setRowCount(len(self.df))
        table.setColumnCount(len(self.df.columns))
        
        # Set headers
        table.setHorizontalHeaderLabels(self.df.columns)
        
        # Fill data
        for i in range(len(self.df)):
            for j in range(len(self.df.columns)):
                item = QTableWidgetItem(str(self.df.iloc[i, j]))
                table.setItem(i, j, item)
        
        layout.addWidget(table)
        
        # Add close button
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
        
        self.setLayout(layout)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df = None
        self.setup_ui()
    
    def setup_ui(self):
        self.setWindowTitle("DOCX Table Converter")
        self.setMinimumSize(800, 600)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Input section
        input_group = QGroupBox("Input")
        input_layout = QVBoxLayout()
        
        # File input
        file_layout = QHBoxLayout()
        self.file_path = QLineEdit()
        self.file_path.setReadOnly(True)
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(QLabel("Input File:"))
        file_layout.addWidget(self.file_path)
        file_layout.addWidget(browse_btn)
        input_layout.addLayout(file_layout)
        
        # Sheet selector (initially hidden)
        self.sheet_widget = QWidget()
        self.sheet_layout = QHBoxLayout(self.sheet_widget)
        self.sheet_label = QLabel("Sheet:")
        self.sheet_combo = QComboBox()
        self.sheet_layout.addWidget(self.sheet_label)
        self.sheet_layout.addWidget(self.sheet_combo)
        self.sheet_layout.addStretch()
        input_layout.addWidget(self.sheet_widget)
        self.sheet_widget.setVisible(False)
        
        # Paste button
        paste_btn = QPushButton("Paste Table Data")
        paste_btn.clicked.connect(self.show_paste_dialog)
        input_layout.addWidget(paste_btn)
        
        # Header rows
        header_layout = QHBoxLayout()
        self.header_rows = QSpinBox()
        self.header_rows.setMinimum(1)
        self.header_rows.setMaximum(10)
        header_layout.addWidget(QLabel("Header Rows:"))
        header_layout.addWidget(self.header_rows)
        header_layout.addStretch()
        input_layout.addLayout(header_layout)
        
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)
        
        # Formatting section
        format_group = QGroupBox("Formatting")
        format_layout = QFormLayout()
        
        # Font settings
        self.font_name = QComboBox()
        self.font_name.addItems(['Times New Roman', 'Arial', 'Calibri', 'SimSun'])
        format_layout.addRow("Font:", self.font_name)
        
        self.font_size = QSpinBox()
        self.font_size.setRange(8, 72)
        self.font_size.setValue(12)
        format_layout.addRow("Font Size:", self.font_size)
        
        # Caption settings
        self.caption_text = QLineEdit("Table")
        format_layout.addRow("Caption:", self.caption_text)
        
        self.caption_desc = QLineEdit()
        format_layout.addRow("Description:", self.caption_desc)
        
        self.caption_bold = QCheckBox("Bold Caption")
        self.caption_bold.setChecked(True)
        format_layout.addRow("", self.caption_bold)
        
        format_group.setLayout(format_layout)
        layout.addWidget(format_group)
        
        # Preview and Export section
        button_layout = QHBoxLayout()
        
        preview_btn = QPushButton("Preview")
        preview_btn.clicked.connect(self.preview_table)
        button_layout.addWidget(preview_btn)
        
        export_btn = QPushButton("Export")
        export_btn.clicked.connect(self.export_table)
        button_layout.addWidget(export_btn)
        
        layout.addLayout(button_layout)
    
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Input File",
            "",
            "Table Files (*.csv *.xls *.xlsx);;All Files (*.*)"
        )
        
        if file_path:
            self.file_path.setText(file_path)
            file_ext = os.path.splitext(file_path)[1].lower()
            
            # Show/hide sheet selector based on file type
            is_excel = file_ext in ['.xls', '.xlsx']
            self.sheet_widget.setVisible(is_excel)
            
            if is_excel:
                # Get available sheets
                try:
                    xls = pd.ExcelFile(file_path)
                    self.sheet_combo.clear()
                    self.sheet_combo.addItems(xls.sheet_names)
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Error reading Excel file: {str(e)}")
                    return
            
            try:
                # Get selected sheet name if it's an Excel file
                sheet_name = self.sheet_combo.currentText() if is_excel else None
                
                self.df = read_table_from_file(
                    file_path,
                    sheet_name=sheet_name,
                    header_rows=self.header_rows.value()
                )
                QMessageBox.information(self, "Success", "File loaded successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error loading file: {str(e)}")
    
    def show_paste_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Paste Table Data")
        dialog.setMinimumSize(600, 400)
        
        layout = QVBoxLayout()
        
        # Text edit for pasting
        text_edit = QTextEdit()
        layout.addWidget(text_edit)
        
        # Buttons
        button_layout = QHBoxLayout()
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        button_layout.addWidget(cancel_btn)
        
        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(dialog.accept)
        button_layout.addWidget(ok_btn)
        
        layout.addLayout(button_layout)
        dialog.setLayout(layout)
        
        if dialog.exec_() == QDialog.Accepted:
            try:
                text = text_edit.toPlainText()
                self.df = parse_clipboard_data(
                    text,
                    header_rows=self.header_rows.value()
                )
                QMessageBox.information(self, "Success", "Data parsed successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error parsing data: {str(e)}")
    
    def preview_table(self):
        if self.df is None:
            QMessageBox.warning(self, "Warning", "Please load or paste table data first!")
            return
        
        dialog = TablePreviewDialog(self.df, self)
        dialog.exec_()
    
    def export_table(self):
        if self.df is None:
            QMessageBox.warning(self, "Warning", "Please load or paste table data first!")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save DOCX File",
            "",
            "Word Documents (*.docx);;All Files (*.*)"
        )
        
        if file_path:
            try:
                write_table_to_docx(
                    df=self.df,
                    output_path=file_path,
                    table_caption=self.caption_text.text(),
                    table_description=self.caption_desc.text(),
                    font_name=self.font_name.currentText(),
                    font_size=self.font_size.value(),
                    caption_bold=self.caption_bold.isChecked(),
                    special_formatting=True
                )
                QMessageBox.information(self, "Success", "Table exported successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error exporting table: {str(e)}")


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 