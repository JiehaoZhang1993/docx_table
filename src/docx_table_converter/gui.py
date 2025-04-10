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
        self.current_language = "zh"  # Default language set to Chinese
        self.setup_ui()
    
    def setup_ui(self):
        self.setWindowTitle("DOCX Table Converter")
        self.setMinimumSize(800, 600)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Language selection
        lang_layout = QHBoxLayout()
        lang_label = QLabel("Language / 语言:")
        self.lang_combo = QComboBox()
        self.lang_combo.addItems(["中文", "English"])
        self.lang_combo.currentTextChanged.connect(self.change_language)
        lang_layout.addWidget(lang_label)
        lang_layout.addWidget(self.lang_combo)
        lang_layout.addStretch()
        layout.addLayout(lang_layout)
        
        # Input section
        input_group = QGroupBox(self.tr("Input"))
        input_layout = QVBoxLayout()
        
        # File input
        file_layout = QHBoxLayout()
        self.file_path = QLineEdit()
        self.file_path.setReadOnly(True)
        browse_btn = QPushButton(self.tr("Browse"))
        browse_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(QLabel(self.tr("Input File:")))
        file_layout.addWidget(self.file_path)
        file_layout.addWidget(browse_btn)
        input_layout.addLayout(file_layout)
        
        # Sheet selector (initially hidden)
        self.sheet_widget = QWidget()
        self.sheet_layout = QHBoxLayout(self.sheet_widget)
        self.sheet_label = QLabel(self.tr("Sheet:"))
        self.sheet_combo = QComboBox()
        self.sheet_layout.addWidget(self.sheet_label)
        self.sheet_layout.addWidget(self.sheet_combo)
        self.sheet_layout.addStretch()
        input_layout.addWidget(self.sheet_widget)
        self.sheet_widget.setVisible(False)
        
        # Paste button
        paste_btn = QPushButton(self.tr("Paste Table Data"))
        paste_btn.clicked.connect(self.show_paste_dialog)
        input_layout.addWidget(paste_btn)
        
        # Header rows
        header_layout = QHBoxLayout()
        self.header_rows = QSpinBox()
        self.header_rows.setMinimum(1)
        self.header_rows.setMaximum(10)
        header_layout.addWidget(QLabel(self.tr("Header Rows:")))
        header_layout.addWidget(self.header_rows)
        header_layout.addStretch()
        input_layout.addLayout(header_layout)
        
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)
        
        # Formatting section
        format_group = QGroupBox(self.tr("Formatting"))
        format_layout = QFormLayout()
        
        # Font settings
        self.font_name = QComboBox()
        self.font_name.addItems(['Times New Roman', 'Arial', 'Calibri', 'SimSun'])
        format_layout.addRow(self.tr("Font:"), self.font_name)
        
        self.font_size = QSpinBox()
        self.font_size.setRange(8, 72)
        self.font_size.setValue(12)
        format_layout.addRow(self.tr("Font Size:"), self.font_size)
        
        # Caption settings
        self.caption_text = QLineEdit(self.tr("Table"))
        format_layout.addRow(self.tr("Caption:"), self.caption_text)
        
        self.caption_desc = QLineEdit()
        format_layout.addRow(self.tr("Description:"), self.caption_desc)
        
        # Export options
        self.include_index = QCheckBox(self.tr("Include Index"))
        self.include_index.setChecked(False)
        format_layout.addRow("", self.include_index)
        
        # Mode selection
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["append", "overwrite"])
        format_layout.addRow(self.tr("Mode:"), self.mode_combo)
        
        format_group.setLayout(format_layout)
        layout.addWidget(format_group)
        
        # Preview and Export section
        button_layout = QHBoxLayout()
        
        preview_btn = QPushButton(self.tr("Preview"))
        preview_btn.clicked.connect(self.preview_table)
        button_layout.addWidget(preview_btn)
        
        export_btn = QPushButton(self.tr("Export"))
        export_btn.clicked.connect(self.export_table)
        button_layout.addWidget(export_btn)
        
        layout.addLayout(button_layout)
    
    def change_language(self, language):
        """Change the application language."""
        self.current_language = "en" if language == "English" else "zh"
        self.retranslate_ui()
    
    def retranslate_ui(self):
        """Update all UI text based on the current language."""
        # Update window title
        self.setWindowTitle(self.tr("DOCX Table Converter"))
        
        # Update all labels and buttons
        for widget in self.findChildren(QLabel):
            if widget.text():
                widget.setText(self.tr(widget.text()))
        
        for widget in self.findChildren(QPushButton):
            if widget.text():
                widget.setText(self.tr(widget.text()))
        
        for widget in self.findChildren(QGroupBox):
            if widget.title():
                widget.setTitle(self.tr(widget.title()))
        
        for widget in self.findChildren(QCheckBox):
            if widget.text():
                widget.setText(self.tr(widget.text()))
    
    def tr(self, text):
        """Translate text based on current language."""
        translations = {
            "en": {
                "DOCX Table Converter": "DOCX Table Converter",
                "Input": "Input",
                "Input File:": "Input File:",
                "Browse": "Browse",
                "Sheet:": "Sheet:",
                "Paste Table Data": "Paste Table Data",
                "Header Rows:": "Header Rows:",
                "Formatting": "Formatting",
                "Font:": "Font:",
                "Font Size:": "Font Size:",
                "Caption:": "Caption:",
                "Description:": "Description:",
                "Include Index": "Include Index",
                "Mode:": "Mode:",
                "Preview": "Preview",
                "Export": "Export",
                "Table": "Table",
                "Language / 语言:": "Language / 语言:"
            },
            "zh": {
                "DOCX Table Converter": "DOCX表格转换器",
                "Input": "输入",
                "Input File:": "输入文件:",
                "Browse": "浏览",
                "Sheet:": "工作表:",
                "Paste Table Data": "粘贴表格数据",
                "Header Rows:": "表头行数:",
                "Formatting": "格式设置",
                "Font:": "字体:",
                "Font Size:": "字号:",
                "Caption:": "标题:",
                "Description:": "描述:",
                "Include Index": "包含索引",
                "Mode:": "模式:",
                "Preview": "预览",
                "Export": "导出",
                "Table": "表格",
                "Language / 语言:": "Language / 语言:"
            }
        }
        return translations.get(self.current_language, {}).get(text, text)
    
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
                    include_index=self.include_index.isChecked(),
                    mode=self.mode_combo.currentText()
                )
                QMessageBox.information(self, "Success", "Table exported successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error exporting table: {str(e)}")


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 