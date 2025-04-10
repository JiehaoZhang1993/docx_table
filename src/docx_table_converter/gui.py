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
                            QDialog, QLineEdit, QFormLayout, QHeaderView, QSplitter)
from PyQt5.QtCore import Qt, QSize, QTimer
from PyQt5.QtGui import QFont, QIcon
from .core import write_table_to_docx, read_table_from_file, parse_clipboard_data
from docx import Document
import tempfile
import io


class TablePreviewDialog(QDialog):
    def __init__(self, df, parent=None, font_name="Times New Roman", font_size=12, 
                 table_caption="Table", table_description="", include_index=False):
        super().__init__(parent)
        self.df = df
        self.font_name = font_name
        self.font_size = font_size
        self.table_caption = table_caption
        self.table_description = table_description
        self.include_index = include_index
        self.temp_docx_path = None
        self.setWindowTitle("表格预览")
        self.setMinimumSize(1000, 800)
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # 创建分割器，左侧显示数据，右侧显示预览效果
        splitter = QSplitter(Qt.Horizontal)
        
        # 左侧：数据表格
        data_group = QGroupBox("原始数据")
        data_layout = QVBoxLayout()
        
        # 创建表格控件
        table = QTableWidget()
        table.setRowCount(len(self.df))
        table.setColumnCount(len(self.df.columns))
        
        # 设置表头 - 修复类型错误
        if isinstance(self.df.columns, pd.RangeIndex):
            headers = [str(i) for i in self.df.columns]
        else:
            headers = [str(col) for col in self.df.columns]
        table.setHorizontalHeaderLabels(headers)
        
        # 填充数据
        for i in range(len(self.df)):
            for j in range(len(self.df.columns)):
                item = QTableWidgetItem(str(self.df.iloc[i, j]))
                table.setItem(i, j, item)
        
        # 调整列宽
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        data_layout.addWidget(table)
        data_group.setLayout(data_layout)
        
        # 右侧：预览效果
        preview_group = QGroupBox("预览效果")
        preview_layout = QVBoxLayout()
        
        # 创建临时DOCX文件
        temp_docx = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
        temp_docx.close()
        self.temp_docx_path = temp_docx.name
        
        try:
            # 写入DOCX文件
            write_table_to_docx(
                df=self.df,
                output_path=self.temp_docx_path,
                table_caption=self.table_caption,
                table_description=self.table_description,
                font_name=self.font_name,
                font_size=self.font_size,
                include_index=self.include_index,
                mode='overwrite'
            )
            
            # 显示预览信息
            preview_info = QLabel(f"已生成预览文档: {self.temp_docx_path}\n\n"
                                 f"标题: {self.table_caption}\n"
                                 f"描述: {self.table_description}\n"
                                 f"字体: {self.font_name}\n"
                                 f"字号: {self.font_size}\n"
                                 f"包含索引: {'是' if self.include_index else '否'}")
            preview_info.setAlignment(Qt.AlignCenter)
            preview_layout.addWidget(preview_info)
            
            # 按钮布局
            button_layout = QHBoxLayout()
            
            # 添加打开文档按钮
            open_btn = QPushButton("打开预览文档")
            open_btn.clicked.connect(self.open_preview_doc)
            button_layout.addWidget(open_btn)
            
            # 添加保存文档按钮
            save_btn = QPushButton("保存文档")
            save_btn.clicked.connect(self.save_document)
            button_layout.addWidget(save_btn)
            
            preview_layout.addLayout(button_layout)
            
        except Exception as e:
            error_label = QLabel(f"生成预览时出错: {str(e)}")
            error_label.setStyleSheet("color: red;")
            preview_layout.addWidget(error_label)
        
        preview_group.setLayout(preview_layout)
        
        # 添加组件到分割器
        splitter.addWidget(data_group)
        splitter.addWidget(preview_group)
        splitter.setSizes([400, 600])  # 设置初始分割比例
        
        layout.addWidget(splitter)
        
        # 添加关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
        
        self.setLayout(layout)
    
    def open_preview_doc(self):
        """打开预览文档"""
        if self.temp_docx_path and os.path.exists(self.temp_docx_path):
            if sys.platform == 'win32':
                os.startfile(self.temp_docx_path)
            else:
                os.system(f'open "{self.temp_docx_path}"')
    
    def save_document(self):
        """保存文档到用户指定位置"""
        if not self.temp_docx_path or not os.path.exists(self.temp_docx_path):
            QMessageBox.warning(self, "警告", "预览文档不存在，无法保存！")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存DOCX文件",
            "",
            "Word文档 (*.docx);;所有文件 (*.*)"
        )
        
        if file_path:
            try:
                # 复制临时文件到用户指定位置
                import shutil
                shutil.copy2(self.temp_docx_path, file_path)
                QMessageBox.information(self, "成功", f"文档已成功保存到: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存文档时出错: {str(e)}")


class PasteTableDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.df = None
        self.raw_df = None  # 存储原始数据
        self.setWindowTitle("粘贴表格数据")
        self.setMinimumSize(800, 600)
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # 说明标签
        instruction_label = QLabel("请从Excel或其他表格软件复制数据，然后点击'粘贴'按钮。")
        layout.addWidget(instruction_label)
        
        # 状态标签（用于显示临时提示）
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("color: green;")
        layout.addWidget(self.status_label)
        
        # 表格控件
        self.table = QTableWidget()
        self.table.setRowCount(10)  # 默认行数
        self.table.setColumnCount(5)  # 默认列数
        
        # 设置表头
        headers = [f"列 {i+1}" for i in range(5)]
        self.table.setHorizontalHeaderLabels(headers)
        
        # 调整列宽
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        layout.addWidget(self.table)
        
        # 表头行数选择
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("表头行数:"))
        self.header_rows = QSpinBox()
        self.header_rows.setMinimum(0)
        self.header_rows.setMaximum(2)
        self.header_rows.setValue(1)
        self.header_rows.valueChanged.connect(self.update_table_headers)
        header_layout.addWidget(self.header_rows)
        header_layout.addStretch()
        layout.addLayout(header_layout)
        
        # 索引列选择
        index_layout = QHBoxLayout()
        index_layout.addWidget(QLabel("索引列:"))
        self.index_column = QComboBox()
        self.index_column.addItem("无")
        # 让下拉框根据内容调整宽度
        self.index_column.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.index_column.currentIndexChanged.connect(self.update_dataframe)
        index_layout.addWidget(self.index_column)
        index_layout.addStretch()
        layout.addLayout(index_layout)
        
        # 按钮布局
        button_layout = QHBoxLayout()
        
        # 粘贴按钮
        paste_btn = QPushButton("粘贴")
        paste_btn.clicked.connect(self.paste_data)
        button_layout.addWidget(paste_btn)
        
        # 清除按钮
        clear_btn = QPushButton("清除")
        clear_btn.clicked.connect(self.clear_table)
        button_layout.addWidget(clear_btn)
        
        # 确定按钮
        ok_btn = QPushButton("确定")
        ok_btn.clicked.connect(self.accept)
        button_layout.addWidget(ok_btn)
        
        # 取消按钮
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
    
    def show_status_message(self, message, duration=3000):
        """显示临时状态消息"""
        self.status_label.setText(message)
        QTimer.singleShot(duration, lambda: self.status_label.setText(""))
    
    def paste_data(self):
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        
        if not text:
            self.show_status_message("剪贴板为空！", 3000)
            return
        
        # 解析剪贴板数据
        try:
            # 尝试检测分隔符
            if '\t' in text:
                delimiter = '\t'
            else:
                delimiter = ','
            
            # 转换为DataFrame
            buffer = io.StringIO(text)
            self.raw_df = pd.read_csv(buffer, sep=delimiter, header=None)
            
            # 更新表格
            self.table.setRowCount(len(self.raw_df))
            self.table.setColumnCount(len(self.raw_df.columns))
            
            # 填充数据
            for i in range(len(self.raw_df)):
                for j in range(len(self.raw_df.columns)):
                    item = QTableWidgetItem(str(self.raw_df.iloc[i, j]))
                    self.table.setItem(i, j, item)
            
            # 更新索引列下拉框
            self.update_index_column_combo()
            
            # 更新表头
            self.update_table_headers()
            
            # 更新DataFrame
            self.update_dataframe()
            
            # 显示临时成功消息
            self.show_status_message("数据已成功粘贴！", 3000)
        except Exception as e:
            self.show_status_message(f"解析数据时出错: {str(e)}", 3000)
    
    def update_index_column_combo(self):
        """更新索引列下拉框"""
        if self.raw_df is None:
            return
        
        self.index_column.clear()
        self.index_column.addItem("无")
        
        # 使用解析后的列名（如果存在）或原始列位置填充
        header_rows = self.header_rows.value()
        if header_rows > 0 and len(self.raw_df) >= header_rows:
            if header_rows == 1:
                col_names = [str(c) for c in self.raw_df.iloc[0]]
            else: # header_rows == 2
                 col_names = [f"{self.raw_df.iloc[0, i]} - {self.raw_df.iloc[1, i]}" for i in range(len(self.raw_df.columns))]
            for name in col_names:
                 self.index_column.addItem(name)
        else:
            # 如果没有表头或行数不足，使用默认列名
            for i in range(len(self.raw_df.columns)):
                col_name = f"列 {i+1}"
                self.index_column.addItem(col_name)
    
    def update_table_headers(self):
        """根据表头行数更新表格表头"""
        if self.raw_df is None:
            return
        
        header_rows = self.header_rows.value()
        
        if header_rows == 0:
            # 无表头
            headers = [f"列 {i+1}" for i in range(len(self.raw_df.columns))]
            self.table.setHorizontalHeaderLabels(headers)
        elif header_rows == 1:
            # 单行表头
            if len(self.raw_df) > 0:
                headers = [str(self.raw_df.iloc[0, i]) for i in range(len(self.raw_df.columns))]
                self.table.setHorizontalHeaderLabels(headers)
            else:
                headers = [f"列 {i+1}" for i in range(len(self.raw_df.columns))]
                self.table.setHorizontalHeaderLabels(headers)
        elif header_rows == 2:
            # 双行表头
            if len(self.raw_df) > 1:
                headers = [f"{self.raw_df.iloc[0, i]} - {self.raw_df.iloc[1, i]}" 
                          for i in range(len(self.raw_df.columns))]
                self.table.setHorizontalHeaderLabels(headers)
            elif len(self.raw_df) > 0:
                headers = [str(self.raw_df.iloc[0, i]) for i in range(len(self.raw_df.columns))]
                self.table.setHorizontalHeaderLabels(headers)
            else:
                headers = [f"列 {i+1}" for i in range(len(self.raw_df.columns))]
                self.table.setHorizontalHeaderLabels(headers)
    
    def update_dataframe(self):
        """根据表头行数和索引列更新DataFrame"""
        if self.raw_df is None:
            return
        
        header_rows = self.header_rows.value()
        index_col = self.index_column.currentIndex() - 1  # -1 表示无索引列
        
        # 创建新的DataFrame
        if header_rows == 0:
            # 无表头
            self.df = self.raw_df.copy()
            # 如果没有表头，列名设为 列1, 列2...
            self.df.columns = [f"列 {i+1}" for i in range(len(self.df.columns))]
        elif header_rows == 1:
            # 单行表头
            if len(self.raw_df) > 0:
                self.df = self.raw_df.iloc[1:].copy()
                self.df.columns = self.raw_df.iloc[0]
            else:
                self.df = self.raw_df.copy()
        elif header_rows == 2:
            # 双行表头
            if len(self.raw_df) > 1:
                self.df = self.raw_df.iloc[2:].copy()
                # 合并两行表头
                self.df.columns = [f"{self.raw_df.iloc[0, i]} - {self.raw_df.iloc[1, i]}" 
                                  for i in range(len(self.raw_df.columns))]
            elif len(self.raw_df) > 0:
                self.df = self.raw_df.iloc[1:].copy()
                self.df.columns = self.raw_df.iloc[0]
            else:
                self.df = self.raw_df.copy()
        
        # 设置索引列
        if index_col >= 0 and index_col < len(self.df.columns):
            self.df.set_index(self.df.columns[index_col], inplace=True)
    
    def clear_table(self):
        self.table.clearContents()
        self.raw_df = None
        self.df = None
        self.index_column.clear()
        self.index_column.addItem("无")
    
    def get_dataframe(self):
        # 应用表头设置
        self.update_dataframe() 
        
        df_copy = self.df.copy() if self.df is not None else None
        if df_copy is None:
            return None
            
        # 根据对话框中的选择设置索引（如果选择了）
        index_col_index = self.index_column.currentIndex()
        if index_col_index > 0: # 0是"无"
            col_name_to_set_index = self.index_column.itemText(index_col_index)
            if col_name_to_set_index in df_copy.columns:
                df_copy.set_index(col_name_to_set_index, inplace=True)
            else:
                # 如果列名不匹配（理论上不应发生，除非手动编辑），则忽略
                pass
                
        return df_copy


class DataSourceDialog(QDialog):
    def __init__(self, file_path, sheet_name, is_excel, header_rows, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.is_excel = is_excel
        self.header_rows = header_rows
        self.df = None
        self.raw_df = None
        self.setWindowTitle("数据源选择")
        self.setMinimumSize(800, 600)
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # 说明标签
        instruction_label = QLabel("您同时选择了文件和粘贴了数据，请选择使用哪一个数据源:")
        layout.addWidget(instruction_label)
        
        # 状态标签（用于显示临时提示）
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("color: green;")
        layout.addWidget(self.status_label)
        
        # 按钮布局
        button_layout = QHBoxLayout()
        
        # 文件数据按钮
        file_btn = QPushButton("使用文件数据")
        file_btn.clicked.connect(self.load_file_data)
        button_layout.addWidget(file_btn)
        
        # 粘贴数据按钮
        paste_btn = QPushButton("使用粘贴数据")
        paste_btn.clicked.connect(self.use_pasted_data)
        button_layout.addWidget(paste_btn)
        
        # 取消按钮
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        layout.addLayout(button_layout)
        
        # 表格控件
        self.table = QTableWidget()
        self.table.setRowCount(10)  # 默认行数
        self.table.setColumnCount(5)  # 默认列数
        
        # 设置表头
        headers = [f"列 {i+1}" for i in range(5)]
        self.table.setHorizontalHeaderLabels(headers)
        
        # 调整列宽
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        layout.addWidget(self.table)
        
        # 表头行数选择
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("表头行数:"))
        self.header_rows_spin = QSpinBox()
        self.header_rows_spin.setMinimum(0)
        self.header_rows_spin.setMaximum(2)
        self.header_rows_spin.setValue(self.header_rows)
        self.header_rows_spin.valueChanged.connect(self.update_table_headers)
        header_layout.addWidget(self.header_rows_spin)
        header_layout.addStretch()
        layout.addLayout(header_layout)
        
        # 索引列选择
        index_layout = QHBoxLayout()
        index_layout.addWidget(QLabel("索引列:"))
        self.index_column = QComboBox()
        self.index_column.addItem("无")
        self.index_column.currentIndexChanged.connect(self.update_dataframe)
        index_layout.addWidget(QLabel(self.tr("索引列:")))
        index_layout.addWidget(self.index_column)
        options_layout.addLayout(index_layout)
        
        # 添加弹性空间
        options_layout.addStretch()
        
        input_layout.addLayout(options_layout)
        
        # 粘贴按钮
        paste_btn = QPushButton(self.tr("粘贴表格数据"))
        paste_btn.clicked.connect(self.show_paste_dialog)
        input_layout.addWidget(paste_btn)
        
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)
        
        # 格式部分
        format_group = QGroupBox(self.tr("格式设置"))
        format_layout = QFormLayout()
        
        # 字体设置
        self.font_name = QComboBox()
        self.font_name.addItems(['宋体', 'Times New Roman', 'Arial', 'Calibri', 'SimSun'])
        format_layout.addRow(self.tr("字体:"), self.font_name)
        
        self.font_size = QSpinBox()
        self.font_size.setRange(8, 72)
        self.font_size.setValue(12)
        format_layout.addRow(self.tr("字号:"), self.font_size)
        
        # 标题设置
        self.caption_text = QLineEdit(self.tr("表格"))
        format_layout.addRow(self.tr("标题:"), self.caption_text)
        
        self.caption_desc = QLineEdit()
        format_layout.addRow(self.tr("描述:"), self.caption_desc)
        
        # 导出选项
        self.include_index = QCheckBox(self.tr("包含索引"))
        self.include_index.setChecked(False)
        format_layout.addRow("", self.include_index)
        
        # 模式选择
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["append", "overwrite"])
        format_layout.addRow(self.tr("模式:"), self.mode_combo)
        
        format_group.setLayout(format_layout)
        layout.addWidget(format_group)
        
        # 预览和导出部分
        button_layout = QHBoxLayout()
        
        preview_btn = QPushButton(self.tr("预览"))
        preview_btn.clicked.connect(self.preview_table)
        button_layout.addWidget(preview_btn)
        
        export_btn = QPushButton(self.tr("导出"))
        export_btn.clicked.connect(self.export_table)
        button_layout.addWidget(export_btn)
        
        layout.addLayout(button_layout)
    
    def change_language(self, language):
        """更改应用程序语言"""
        self.current_language = "en" if language == "English" else "zh"
        self.retranslate_ui()
    
    def retranslate_ui(self):
        """根据当前语言更新所有UI文本"""
        # 更新窗口标题
        self.setWindowTitle(self.tr("论文三线表--一键转换"))
        
        # 更新所有标签和按钮
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
        """根据当前语言翻译文本"""
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
                "Language / 语言:": "Language / 语言:",
                "输入": "Input",
                "输入文件:": "Input File:",
                "浏览": "Browse",
                "工作表:": "Sheet:",
                "粘贴表格数据": "Paste Table Data",
                "表头行数:": "Header Rows:",
                "格式设置": "Formatting",
                "字体:": "Font:",
                "字号:": "Font Size:",
                "标题:": "Caption:",
                "描述:": "Description:",
                "包含索引": "Include Index",
                "模式:": "Mode:",
                "预览": "Preview",
                "导出": "Export",
                "表格": "Table",
                "论文三线表--一键转换": "Research Table Converter",
                "数据源选择": "Data Source Selection",
                "您同时选择了文件和粘贴了数据，请选择使用哪一个数据源:": "You have both selected a file and pasted data. Please choose which data source to use:",
                "使用文件数据": "Use File Data",
                "使用粘贴数据": "Use Pasted Data",
                "取消": "Cancel",
                "索引列:": "Index Column:",
                "无": "None",
                "导入成功": "Import Success"
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
                "Language / 语言:": "Language / 语言:",
                "输入": "输入",
                "输入文件:": "输入文件:",
                "浏览": "浏览",
                "工作表:": "工作表:",
                "粘贴表格数据": "粘贴表格数据",
                "表头行数:": "表头行数:",
                "格式设置": "格式设置",
                "字体:": "字体:",
                "字号:": "字号:",
                "标题:": "标题:",
                "描述:": "描述:",
                "包含索引": "包含索引",
                "模式:": "模式:",
                "预览": "预览",
                "导出": "导出",
                "表格": "表格",
                "论文三线表--一键转换": "论文三线表--一键转换",
                "数据源选择": "数据源选择",
                "您同时选择了文件和粘贴了数据，请选择使用哪一个数据源:": "您同时选择了文件和粘贴了数据，请选择使用哪一个数据源:",
                "使用文件数据": "使用文件数据",
                "使用粘贴数据": "使用粘贴数据",
                "取消": "取消",
                "索引列:": "索引列:",
                "无": "无",
                "导入成功": "导入成功"
            }
        }
        return translations.get(self.current_language, {}).get(text, text)
    
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择输入文件",
            "",
            "表格文件 (*.csv *.xls *.xlsx);;所有文件 (*.*)"
        )
        
        if file_path:
            self.file_path.setText(file_path)
            file_ext = os.path.splitext(file_path)[1].lower()
            
            # 根据文件类型显示/隐藏工作表选择器
            is_excel = file_ext in ['.xls', '.xlsx']
            self.sheet_widget.setVisible(is_excel)
            
            if is_excel:
                # 获取可用工作表
                try:
                    xls = pd.ExcelFile(file_path)
                    self.sheet_combo.clear()
                    self.sheet_combo.addItems(xls.sheet_names)
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"读取Excel文件时出错: {str(e)}")
                    return
            
            # 获取选定的工作表名称（如果是Excel文件）
            sheet_name = self.sheet_combo.currentText() if is_excel else None
            header_rows = self.header_rows.value()
            
            # 如果已经有粘贴的数据，询问用户使用哪个数据源
            if self.df is not None:
                dialog = DataSourceDialog(
                    file_path, 
                    sheet_name, 
                    is_excel, 
                    header_rows,
                    self
                )
                
                if dialog.exec_() == QDialog.Accepted:
                    self.df = dialog.get_dataframe()
                    self.header_rows.setValue(dialog.header_rows_spin.value())
                    self.update_index_column_combo() # 更新主窗口的索引列
                    self.import_status.setText(self.tr("导入成功"))
                    QTimer.singleShot(3000, lambda: self.import_status.setText(""))
            else:
                # 直接加载文件数据
                self.load_file_data(file_path, sheet_name, is_excel, header_rows)
    
    def load_file_data(self, file_path, sheet_name, is_excel, header_rows):
        """加载文件数据"""
        try:
            # 直接使用 core 函数读取，传入正确的 header_rows
            self.df = read_table_from_file(
                file_path,
                sheet_name=sheet_name,
                header_rows=header_rows
            )
            
            # 更新索引列下拉框
            self.update_index_column_combo()
            
            # 显示导入成功状态
            self.import_status.setText(self.tr("导入成功"))
            QTimer.singleShot(3000, lambda: self.import_status.setText(""))
        except Exception as e:
            # 在错误信息中包含 sheet_name (如果存在)
            error_msg = f"加载文件数据时出错: {str(e)}"
            if is_excel and sheet_name:
                error_msg += f" (sheet: {sheet_name})"
            QMessageBox.critical(self, "错误", error_msg)
            self.df = None # 清空DataFrame防止后续操作出错
            self.update_index_column_combo() # 清空索引列下拉框
            self.import_status.setText("") # 清空状态显示
    
    def update_index_column_combo(self):
        """更新主窗口的索引列下拉框"""
        self.index_column.clear()
        self.index_column.addItem(self.tr("无"))
        
        if self.df is not None and not self.df.empty:
            for i, col in enumerate(self.df.columns):
                self.index_column.addItem(str(col))
            
            # 如果DataFrame有索引名且包含索引，尝试自动选择
            if self.include_index.isChecked() and self.df.index.name is not None:
                try:
                    # 查找索引名在列名中的位置
                    idx_pos = list(self.df.columns).index(self.df.index.name)
                    self.index_column.setCurrentIndex(idx_pos + 1) # +1 因为第一项是 "无"
                except ValueError:
                    pass # 索引名不在列中，保持默认 "无"
            elif not isinstance(self.df.index, pd.RangeIndex):
                 # 如果有非默认索引但未勾选包含，或者没有索引名，也默认选中第一个列作为可能的索引
                 if len(self.df.columns) > 0: # 确保有列可选
                    self.index_column.setCurrentIndex(1)
    
    def show_paste_dialog(self):
        dialog = PasteTableDialog(self)
        
        if dialog.exec_() == QDialog.Accepted:
            try:
                pasted_df = dialog.get_dataframe()
                header_rows = dialog.header_rows.value() # 获取粘贴对话框的表头设置
                
                # 如果已经有文件数据，询问用户使用哪个数据源
                if self.df is not None and self.file_path.text():
                    dialog = DataSourceDialog(
                        self.file_path.text(), 
                        self.sheet_combo.currentText() if self.sheet_widget.isVisible() else None, 
                        self.sheet_widget.isVisible(), 
                        self.header_rows.value(),
                        self
                    )
                    
                    if dialog.exec_() == QDialog.Accepted:
                        self.df = dialog.get_dataframe()
                        self.header_rows.setValue(dialog.header_rows_spin.value())
                        # 如果DataSourceDialog接受，说明用户选了某个源
                        self.update_index_column_combo() # 更新主窗口索引列
                        # 检查并设置包含索引复选框
                        self.include_index.setChecked(not isinstance(self.df.index, pd.RangeIndex))
                    else:
                        # 直接使用粘贴的数据
                        self.df = pasted_df
                        self.header_rows.setValue(header_rows) # 更新主窗口的表头设置
                        self.update_index_column_combo() # 更新主窗口索引列
                        # 检查并设置包含索引复选框 (因为get_dataframe可能已设置索引)
                        if self.df is not None:
                            self.include_index.setChecked(not isinstance(self.df.index, pd.RangeIndex))
                        self.file_path.setText("") # 清空文件路径
                        self.import_status.setText("") # 清空导入状态
                else:
                    # 直接使用粘贴的数据
                    self.df = pasted_df
                    self.header_rows.setValue(header_rows) # 更新主窗口的表头设置
                    self.update_index_column_combo() # 更新主窗口索引列
                    # 检查并设置包含索引复选框 (因为get_dataframe可能已设置索引)
                    if self.df is not None:
                        self.include_index.setChecked(not isinstance(self.df.index, pd.RangeIndex))
                    self.file_path.setText("") # 清空文件路径
                    self.import_status.setText("") # 清空导入状态
            except Exception as e:
                QMessageBox.critical(self, "错误", f"处理粘贴数据时出错: {str(e)}")
                self.df = None # 清空DataFrame
                self.update_index_column_combo() # 清空索引列
    
    def preview_table(self):
        if self.df is None or self.df.empty:
            QMessageBox.warning(self, "警告", "请先加载或粘贴有效的表格数据！")
            return
        
        try:
            # 预览时使用临时的DataFrame副本，避免修改原始df
            preview_df = self.df.copy()
            include_index = self.include_index.isChecked()
            index_col_selection = self.index_column.currentIndex()
            
            # 处理索引列
            if include_index:
                if index_col_selection > 0:
                    # 用户选择了特定列作为索引
                    col_name = self.index_column.itemText(index_col_selection)
                    if col_name in preview_df.columns:
                        preview_df.set_index(col_name, inplace=True)
                    else:
                        # 如果选中的列名无效（可能发生在数据更改后），则不设置索引
                        QMessageBox.warning(self, "警告", f"选择的索引列 '{col_name}' 不存在，将不包含索引。")
                        include_index = False 
                elif not isinstance(preview_df.index, pd.RangeIndex):
                    # 如果用户勾选了包含索引，但下拉框选了"无"，且df本身已有非默认索引，则使用该索引
                    pass # 保持现有索引
                else:
                     # 用户勾选了包含索引，下拉框选了"无"，且df是默认索引，则不包含索引
                     include_index = False
                     QMessageBox.warning(self, "警告", "请选择一个有效的索引列，或取消勾选'包含索引'。")
                     return # 阻止预览
            else:
                 # 如果用户未勾选包含索引，确保重置为默认RangeIndex
                 if not isinstance(preview_df.index, pd.RangeIndex):
                     preview_df.reset_index(inplace=True)
                 # 并且，如果用户在下拉框中选择了某个列，则将其删除
                 if index_col_selection > 0:
                     col_name_to_drop = self.index_column.itemText(index_col_selection)
                     if col_name_to_drop in preview_df.columns:
                         preview_df.drop(columns=[col_name_to_drop], inplace=True)
            
            dialog = TablePreviewDialog(
                preview_df, 
                self,
                font_name=self.font_name.currentText(),
                font_size=self.font_size.value(),
                table_caption=self.caption_text.text(),
                table_description=self.caption_desc.text(),
                include_index=include_index # 使用更新后的值
            )
            dialog.exec_()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"预览表格时出错: {str(e)}")
    
    def export_table(self):
        if self.df is None or self.df.empty:
            QMessageBox.warning(self, "警告", "请先加载或粘贴有效的表格数据！")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存DOCX文件",
            "",
            "Word文档 (*.docx);;所有文件 (*.*)"
        )
        
        if file_path:
            try:
                # 导出时使用临时的DataFrame副本，避免修改原始df
                export_df = self.df.copy()
                include_index = self.include_index.isChecked()
                index_col_selection = self.index_column.currentIndex()
                
                # 处理索引列 (与预览逻辑相同)
                if include_index:
                    if index_col_selection > 0:
                        col_name = self.index_column.itemText(index_col_selection)
                        if col_name in export_df.columns:
                           export_df.set_index(col_name, inplace=True)
                        else:
                           QMessageBox.warning(self, "警告", f"选择的索引列 '{col_name}' 不存在，将不包含索引。")
                           include_index = False
                    elif not isinstance(export_df.index, pd.RangeIndex):
                        pass # 保持现有索引
                    else:
                         include_index = False
                         QMessageBox.warning(self, "警告", "请选择一个有效的索引列，或取消勾选'包含索引'以导出。")
                         return # 阻止导出
                else:
                     if not isinstance(export_df.index, pd.RangeIndex):
                         export_df.reset_index(inplace=True)
                     # 并且，如果用户在下拉框中选择了某个列，则将其删除
                     if index_col_selection > 0:
                         col_name_to_drop = self.index_column.itemText(index_col_selection)
                         if col_name_to_drop in export_df.columns:
                             export_df.drop(columns=[col_name_to_drop], inplace=True)
                
                write_table_to_docx(
                    df=export_df,
                    output_path=file_path,
                    table_caption=self.caption_text.text(),
                    table_description=self.caption_desc.text(),
                    font_name=self.font_name.currentText(),
                    font_size=self.font_size.value(),
                    include_index=include_index, # 使用更新后的值
                    mode=self.mode_combo.currentText()
                )
                QMessageBox.information(self, "成功", "表格导出成功！")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出表格时出错: {str(e)}")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.df = None
        self.current_language = "zh"  # 默认语言设为中文
        self.setup_ui()
    
    def setup_ui(self):
        self.setWindowTitle("论文三线表--一键转换")
        self.setMinimumSize(800, 600)
        
        # 创建中央部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # 语言选择
        lang_layout = QHBoxLayout()
        lang_label = QLabel("Language / 语言:")
        self.lang_combo = QComboBox()
        self.lang_combo.addItems(["中文", "English"])
        self.lang_combo.currentTextChanged.connect(self.change_language)
        lang_layout.addWidget(lang_label)
        lang_layout.addWidget(self.lang_combo)
        lang_layout.addStretch()
        layout.addLayout(lang_layout)
        
        # 输入部分
        input_group = QGroupBox(self.tr("输入"))
        input_layout = QVBoxLayout()
        
        # 文件输入
        file_layout = QHBoxLayout()
        self.file_path = QLineEdit()
        self.file_path.setReadOnly(True)
        browse_btn = QPushButton(self.tr("浏览"))
        browse_btn.clicked.connect(self.browse_file)
        
        # 添加状态标签（用于显示导入成功）
        self.import_status = QLabel("")
        self.import_status.setStyleSheet("color: green;")
        self.import_status.setMinimumWidth(100)
        
        file_layout.addWidget(QLabel(self.tr("输入文件:")))
        file_layout.addWidget(self.file_path)
        file_layout.addWidget(browse_btn)
        file_layout.addWidget(self.import_status)
        input_layout.addLayout(file_layout)
        
        # 工作表选择和表头行数选择（在同一行）
        options_layout = QHBoxLayout()
        
        # 工作表选择（初始隐藏）
        self.sheet_widget = QWidget()
        self.sheet_layout = QHBoxLayout(self.sheet_widget)
        self.sheet_layout.setContentsMargins(0, 0, 0, 0)  # 减少边距
        self.sheet_label = QLabel(self.tr("工作表:"))
        self.sheet_combo = QComboBox()
        self.sheet_layout.addWidget(self.sheet_label)
        self.sheet_layout.addWidget(self.sheet_combo)
        self.sheet_widget.setVisible(False)
        options_layout.addWidget(self.sheet_widget)
        
        # 表头行数选择
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)  # 减少边距
        self.header_rows = QSpinBox()
        self.header_rows.setMinimum(0)
        self.header_rows.setMaximum(2)
        self.header_rows.setValue(1)
        header_layout.addWidget(QLabel(self.tr("表头行数:")))
        header_layout.addWidget(self.header_rows)
        options_layout.addLayout(header_layout)
        
        # 索引列选择
        index_layout = QHBoxLayout()
        index_layout.setContentsMargins(0, 0, 0, 0)  # 减少边距
        self.index_column = QComboBox()
        self.index_column.addItem("无")
        # 让下拉框根据内容调整宽度
        self.index_column.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        index_layout.addWidget(QLabel(self.tr("索引列:")))
        index_layout.addWidget(self.index_column)
        options_layout.addLayout(index_layout)
        
        # 添加弹性空间
        options_layout.addStretch()
        
        input_layout.addLayout(options_layout)
        
        # 粘贴按钮
        paste_btn = QPushButton(self.tr("粘贴表格数据"))
        paste_btn.clicked.connect(self.show_paste_dialog)
        input_layout.addWidget(paste_btn)
        
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)
        
        # 格式部分
        format_group = QGroupBox(self.tr("格式设置"))
        format_layout = QFormLayout()
        
        # 字体设置
        self.font_name = QComboBox()
        self.font_name.addItems(['宋体', 'Times New Roman', 'Arial', 'Calibri', 'SimSun'])
        format_layout.addRow(self.tr("字体:"), self.font_name)
        
        self.font_size = QSpinBox()
        self.font_size.setRange(8, 72)
        self.font_size.setValue(12)
        format_layout.addRow(self.tr("字号:"), self.font_size)
        
        # 标题设置
        self.caption_text = QLineEdit(self.tr("表格"))
        format_layout.addRow(self.tr("标题:"), self.caption_text)
        
        self.caption_desc = QLineEdit()
        format_layout.addRow(self.tr("描述:"), self.caption_desc)
        
        # 导出选项
        self.include_index = QCheckBox(self.tr("包含索引"))
        self.include_index.setChecked(False)
        format_layout.addRow("", self.include_index)
        
        # 模式选择
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["append", "overwrite"])
        format_layout.addRow(self.tr("模式:"), self.mode_combo)
        
        format_group.setLayout(format_layout)
        layout.addWidget(format_group)
        
        # 预览和导出部分
        button_layout = QHBoxLayout()
        
        preview_btn = QPushButton(self.tr("预览"))
        preview_btn.clicked.connect(self.preview_table)
        button_layout.addWidget(preview_btn)
        
        export_btn = QPushButton(self.tr("导出"))
        export_btn.clicked.connect(self.export_table)
        button_layout.addWidget(export_btn)
        
        layout.addLayout(button_layout)
    
    def change_language(self, language):
        """更改应用程序语言"""
        self.current_language = "en" if language == "English" else "zh"
        self.retranslate_ui()
    
    def retranslate_ui(self):
        """根据当前语言更新所有UI文本"""
        # 更新窗口标题
        self.setWindowTitle(self.tr("论文三线表--一键转换"))
        
        # 更新所有标签和按钮
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
        """根据当前语言翻译文本"""
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
                "Language / 语言:": "Language / 语言:",
                "输入": "Input",
                "输入文件:": "Input File:",
                "浏览": "Browse",
                "工作表:": "Sheet:",
                "粘贴表格数据": "Paste Table Data",
                "表头行数:": "Header Rows:",
                "格式设置": "Formatting",
                "字体:": "Font:",
                "字号:": "Font Size:",
                "标题:": "Caption:",
                "描述:": "Description:",
                "包含索引": "Include Index",
                "模式:": "Mode:",
                "预览": "Preview",
                "导出": "Export",
                "表格": "Table",
                "论文三线表--一键转换": "Research Table Converter",
                "数据源选择": "Data Source Selection",
                "您同时选择了文件和粘贴了数据，请选择使用哪一个数据源:": "You have both selected a file and pasted data. Please choose which data source to use:",
                "使用文件数据": "Use File Data",
                "使用粘贴数据": "Use Pasted Data",
                "取消": "Cancel",
                "索引列:": "Index Column:",
                "无": "None",
                "导入成功": "Import Success"
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
                "Language / 语言:": "Language / 语言:",
                "输入": "输入",
                "输入文件:": "输入文件:",
                "浏览": "浏览",
                "工作表:": "工作表:",
                "粘贴表格数据": "粘贴表格数据",
                "表头行数:": "表头行数:",
                "格式设置": "格式设置",
                "字体:": "字体:",
                "字号:": "字号:",
                "标题:": "标题:",
                "描述:": "描述:",
                "包含索引": "包含索引",
                "模式:": "模式:",
                "预览": "预览",
                "导出": "导出",
                "表格": "表格",
                "论文三线表--一键转换": "论文三线表--一键转换",
                "数据源选择": "数据源选择",
                "您同时选择了文件和粘贴了数据，请选择使用哪一个数据源:": "您同时选择了文件和粘贴了数据，请选择使用哪一个数据源:",
                "使用文件数据": "使用文件数据",
                "使用粘贴数据": "使用粘贴数据",
                "取消": "取消",
                "索引列:": "索引列:",
                "无": "无",
                "导入成功": "导入成功"
            }
        }
        return translations.get(self.current_language, {}).get(text, text)
    
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择输入文件",
            "",
            "表格文件 (*.csv *.xls *.xlsx);;所有文件 (*.*)"
        )
        
        if file_path:
            self.file_path.setText(file_path)
            file_ext = os.path.splitext(file_path)[1].lower()
            
            # 根据文件类型显示/隐藏工作表选择器
            is_excel = file_ext in ['.xls', '.xlsx']
            self.sheet_widget.setVisible(is_excel)
            
            if is_excel:
                # 获取可用工作表
                try:
                    xls = pd.ExcelFile(file_path)
                    self.sheet_combo.clear()
                    self.sheet_combo.addItems(xls.sheet_names)
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"读取Excel文件时出错: {str(e)}")
                    return
            
            # 获取选定的工作表名称（如果是Excel文件）
            sheet_name = self.sheet_combo.currentText() if is_excel else None
            header_rows = self.header_rows.value()
            
            # 如果已经有粘贴的数据，询问用户使用哪个数据源
            if self.df is not None:
                dialog = DataSourceDialog(
                    file_path, 
                    sheet_name, 
                    is_excel, 
                    header_rows,
                    self
                )
                
                if dialog.exec_() == QDialog.Accepted:
                    self.df = dialog.get_dataframe()
                    self.header_rows.setValue(dialog.header_rows_spin.value())
                    self.update_index_column_combo() # 更新主窗口的索引列
                    self.import_status.setText(self.tr("导入成功"))
                    QTimer.singleShot(3000, lambda: self.import_status.setText(""))
            else:
                # 直接加载文件数据
                self.load_file_data(file_path, sheet_name, is_excel, header_rows)
    
    def load_file_data(self, file_path, sheet_name, is_excel, header_rows):
        """加载文件数据"""
        try:
            # 直接使用 core 函数读取，传入正确的 header_rows
            self.df = read_table_from_file(
                file_path,
                sheet_name=sheet_name,
                header_rows=header_rows
            )
            
            # 更新索引列下拉框
            self.update_index_column_combo()
            
            # 显示导入成功状态
            self.import_status.setText(self.tr("导入成功"))
            QTimer.singleShot(3000, lambda: self.import_status.setText(""))
        except Exception as e:
            # 在错误信息中包含 sheet_name (如果存在)
            error_msg = f"加载文件数据时出错: {str(e)}"
            if is_excel and sheet_name:
                error_msg += f" (sheet: {sheet_name})"
            QMessageBox.critical(self, "错误", error_msg)
            self.df = None # 清空DataFrame防止后续操作出错
            self.update_index_column_combo() # 清空索引列下拉框
            self.import_status.setText("") # 清空状态显示
    
    def update_index_column_combo(self):
        """更新主窗口的索引列下拉框"""
        self.index_column.clear()
        self.index_column.addItem(self.tr("无"))
        
        if self.df is not None and not self.df.empty:
            for i, col in enumerate(self.df.columns):
                self.index_column.addItem(str(col))
            
            # 如果DataFrame有索引名且包含索引，尝试自动选择
            if self.include_index.isChecked() and self.df.index.name is not None:
                try:
                    # 查找索引名在列名中的位置
                    idx_pos = list(self.df.columns).index(self.df.index.name)
                    self.index_column.setCurrentIndex(idx_pos + 1) # +1 因为第一项是 "无"
                except ValueError:
                    pass # 索引名不在列中，保持默认 "无"
            elif not isinstance(self.df.index, pd.RangeIndex):
                 # 如果有非默认索引但未勾选包含，或者没有索引名，也默认选中第一个列作为可能的索引
                 if len(self.df.columns) > 0: # 确保有列可选
                    self.index_column.setCurrentIndex(1)
    
    def show_paste_dialog(self):
        dialog = PasteTableDialog(self)
        
        if dialog.exec_() == QDialog.Accepted:
            try:
                pasted_df = dialog.get_dataframe()
                header_rows = dialog.header_rows.value() # 获取粘贴对话框的表头设置
                
                # 如果已经有文件数据，询问用户使用哪个数据源
                if self.df is not None and self.file_path.text():
                    dialog = DataSourceDialog(
                        self.file_path.text(), 
                        self.sheet_combo.currentText() if self.sheet_widget.isVisible() else None, 
                        self.sheet_widget.isVisible(), 
                        self.header_rows.value(),
                        self
                    )
                    
                    if dialog.exec_() == QDialog.Accepted:
                        self.df = dialog.get_dataframe()
                        self.header_rows.setValue(dialog.header_rows_spin.value())
                        # 如果DataSourceDialog接受，说明用户选了某个源
                        self.update_index_column_combo() # 更新主窗口索引列
                        # 检查并设置包含索引复选框
                        self.include_index.setChecked(not isinstance(self.df.index, pd.RangeIndex))
                    else:
                        # 直接使用粘贴的数据
                        self.df = pasted_df
                        self.header_rows.setValue(header_rows) # 更新主窗口的表头设置
                        self.update_index_column_combo() # 更新主窗口索引列
                        # 检查并设置包含索引复选框 (因为get_dataframe可能已设置索引)
                        if self.df is not None:
                            self.include_index.setChecked(not isinstance(self.df.index, pd.RangeIndex))
                        self.file_path.setText("") # 清空文件路径
                        self.import_status.setText("") # 清空导入状态
                else:
                    # 直接使用粘贴的数据
                    self.df = pasted_df
                    self.header_rows.setValue(header_rows) # 更新主窗口的表头设置
                    self.update_index_column_combo() # 更新主窗口索引列
                    # 检查并设置包含索引复选框 (因为get_dataframe可能已设置索引)
                    if self.df is not None:
                        self.include_index.setChecked(not isinstance(self.df.index, pd.RangeIndex))
                    self.file_path.setText("") # 清空文件路径
                    self.import_status.setText("") # 清空导入状态
            except Exception as e:
                QMessageBox.critical(self, "错误", f"处理粘贴数据时出错: {str(e)}")
                self.df = None # 清空DataFrame
                self.update_index_column_combo() # 清空索引列
    
    def preview_table(self):
        if self.df is None or self.df.empty:
            QMessageBox.warning(self, "警告", "请先加载或粘贴有效的表格数据！")
            return
        
        try:
            # 预览时使用临时的DataFrame副本，避免修改原始df
            preview_df = self.df.copy()
            include_index = self.include_index.isChecked()
            index_col_selection = self.index_column.currentIndex()
            
            # 处理索引列
            if include_index:
                if index_col_selection > 0:
                    # 用户选择了特定列作为索引
                    col_name = self.index_column.itemText(index_col_selection)
                    if col_name in preview_df.columns:
                        preview_df.set_index(col_name, inplace=True)
                    else:
                        # 如果选中的列名无效（可能发生在数据更改后），则不设置索引
                        QMessageBox.warning(self, "警告", f"选择的索引列 '{col_name}' 不存在，将不包含索引。")
                        include_index = False 
                elif not isinstance(preview_df.index, pd.RangeIndex):
                    # 如果用户勾选了包含索引，但下拉框选了"无"，且df本身已有非默认索引，则使用该索引
                    pass # 保持现有索引
                else:
                     # 用户勾选了包含索引，下拉框选了"无"，且df是默认索引，则不包含索引
                     include_index = False
                     QMessageBox.warning(self, "警告", "请选择一个有效的索引列，或取消勾选'包含索引'。")
                     return # 阻止预览
            else:
                 # 如果用户未勾选包含索引，确保重置为默认RangeIndex
                 if not isinstance(preview_df.index, pd.RangeIndex):
                     preview_df.reset_index(inplace=True)
                 # 并且，如果用户在下拉框中选择了某个列，则将其删除
                 if index_col_selection > 0:
                     col_name_to_drop = self.index_column.itemText(index_col_selection)
                     if col_name_to_drop in preview_df.columns:
                         preview_df.drop(columns=[col_name_to_drop], inplace=True)
            
            dialog = TablePreviewDialog(
                preview_df, 
                self,
                font_name=self.font_name.currentText(),
                font_size=self.font_size.value(),
                table_caption=self.caption_text.text(),
                table_description=self.caption_desc.text(),
                include_index=include_index # 使用更新后的值
            )
            dialog.exec_()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"预览表格时出错: {str(e)}")
    
    def export_table(self):
        if self.df is None or self.df.empty:
            QMessageBox.warning(self, "警告", "请先加载或粘贴有效的表格数据！")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存DOCX文件",
            "",
            "Word文档 (*.docx);;所有文件 (*.*)"
        )
        
        if file_path:
            try:
                # 导出时使用临时的DataFrame副本，避免修改原始df
                export_df = self.df.copy()
                include_index = self.include_index.isChecked()
                index_col_selection = self.index_column.currentIndex()
                
                # 处理索引列 (与预览逻辑相同)
                if include_index:
                    if index_col_selection > 0:
                        col_name = self.index_column.itemText(index_col_selection)
                        if col_name in export_df.columns:
                           export_df.set_index(col_name, inplace=True)
                        else:
                           QMessageBox.warning(self, "警告", f"选择的索引列 '{col_name}' 不存在，将不包含索引。")
                           include_index = False
                    elif not isinstance(export_df.index, pd.RangeIndex):
                        pass # 保持现有索引
                    else:
                         include_index = False
                         QMessageBox.warning(self, "警告", "请选择一个有效的索引列，或取消勾选'包含索引'以导出。")
                         return # 阻止导出
                else:
                     if not isinstance(export_df.index, pd.RangeIndex):
                         export_df.reset_index(inplace=True)
                     # 并且，如果用户在下拉框中选择了某个列，则将其删除
                     if index_col_selection > 0:
                         col_name_to_drop = self.index_column.itemText(index_col_selection)
                         if col_name_to_drop in export_df.columns:
                             export_df.drop(columns=[col_name_to_drop], inplace=True)
                
                write_table_to_docx(
                    df=export_df,
                    output_path=file_path,
                    table_caption=self.caption_text.text(),
                    table_description=self.caption_desc.text(),
                    font_name=self.font_name.currentText(),
                    font_size=self.font_size.value(),
                    include_index=include_index, # 使用更新后的值
                    mode=self.mode_combo.currentText()
                )
                QMessageBox.information(self, "成功", "表格导出成功！")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出表格时出错: {str(e)}")


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_()) 