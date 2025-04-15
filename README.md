 # DOCX Table Converter
[English](README.md) | [中文](README_zh.md)
A user-friendly tool that helps you convert Excel, CSV, and other table formats into beautifully formatted tables in Word documents. No more hassle with table formatting in your papers and reports!
## Why This Tool?
Have you ever encountered these frustrations:
- Formatting gets messed up when copying tables from Excel to Word
- Need three-line table format for academic papers, but manual adjustment is tedious
- Processing multiple tables is time-consuming
- Special formats (like subscripts) are difficult to display correctly in Word
- Need to batch process multiple tables while maintaining consistent formatting
This tool is created to solve these problems! It can:
- Automatically convert Excel/CSV tables to Word document tables
- Support three-line table format (no vertical lines, only horizontal lines)
- Handle special formats automatically, like subscripts (e.g., mg L$_{-1}$)
- Support multi-level headers for complex data presentation
- Support batch processing of multiple tables
- Provide a user-friendly GUI, no programming knowledge required
## Features
- Import tables from CSV, XLS, XLSX files
- Paste table data directly from clipboard
- Preview tables before export
- Customizable table format:
  - Separate font settings for Chinese and English text
  - Adjustable font size
  - Adjustable table border width
  - Optional bold table caption
- Support multi-level headers (up to 2 levels)
- Support special formats (like subscripts)
- Three-line table format (no vertical lines)
- Bilingual interface (Chinese and English)
- Batch processing features:
  - Import multiple tables simultaneously
  - Option to merge or save separately
  - Individual title and description for each table
  - Preview all tables
- Save functionality:
  - Save directly from preview interface
  - Choose save location
  - Auto-open generated documents
## Installation
The project is not yet published to PyPI. You can install it through:
### Method 1: Clone from GitHub
```bash
git clone https://github.com/JiehaoZhang1993/docx_table.git
cd docx_table
pip install -e .
```
### Method 2: Download Source Code
1. Download and extract the project source code
2. Navigate to the project directory
3. Run `pip install -e .` to install
## Usage
### Method 1: As a Python Package
If you're familiar with Python, you can call it directly in your code:
```python
import pandas as pd
from docx_table_converter import write_table_to_docx, write_tables_to_docx
# Single table processing
df = pd.DataFrame({
    'Sample': ['A', 'B', 'C'],
    'Temperature (°C)': [25, 30, 35],
    'Concentration (mg L$_{-1}$)': [10, 20, 30],
    'pH': [7.0, 7.2, 7.4]
})
# Write single table
write_table_to_docx(
    df=df, 
    output_path='output_table.docx',
    table_caption='Table 1. Sample Data',
    chinese_font='SimSun',
    english_font='Times New Roman',
    font_size=12,
    border_width=1,
    bold_caption=True,
    mode='append'
)
# Batch process multiple tables
tables = [df1, df2, df3]  # List of DataFrames or file paths
captions = ['Table 1', 'Table 2', 'Table 3']
write_tables_to_docx(
    tables=tables,
    captions=captions,
    output_path='batch_tables.docx',
    font_name='Times New Roman',
    font_size=10.5,
    header_rows=[1, 1, 1],  # Number of header rows for each table
    separate_files=False  # True to save each table as separate file
)
```
### Method 2: Using the GUI
If you're not familiar with programming, you can use the graphical interface:
1. Run the program: `python run_gui.py`
2. In the main interface:
   - Click "Browse" to select Excel/CSV files, or "Paste Table Data" to import from clipboard
   - Set header rows (0-2 rows)
   - Customize Chinese and English fonts, font size
   - Set table caption
   - Adjust table border width
   - Choose whether to bold the caption
3. Click "Preview" to check the result, you can save directly from the preview interface
4. Click "Export" to save as Word document
Batch Processing:
1. Click "Batch Process" button
2. In the batch processing interface:
   - Add multiple files or paste multiple tables
   - Set title and description for each table
   - Choose whether to merge or save separately
3. Preview individual or all tables
4. Click "Export" to complete batch conversion
## Parameter Reference
### write_table_to_docx Parameters
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| df | DataFrame | Required | pandas DataFrame to convert |
| output_path | str | Required | Output DOCX file path |
| table_caption | str | "Table 1" | Table caption |
| chinese_font | str | 'SimSun' | Chinese font |
| english_font | str | 'Times New Roman' | English font |
| font_size | int | 12 | Font size |
| header_rows | int | 1 | Number of header rows |
| mode | str | 'append' | File mode: 'append' or 'overwrite' |
| border_width | float | 1 | Table border width |
| bold_caption | bool | True | Whether to bold the table caption |
### write_tables_to_docx Parameters
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| tables | list | Required | List of DataFrames or file paths |
| captions | list | Required | List of table captions |
| output_path | str | Required | Output file path |
| font_name | str | 'Times New Roman' | Font name |
| font_size | float | 10.5 | Font size |
| header_rows | list | None | List of header row counts for each table |
| separate_files | bool | False | Whether to save each table as separate file |
## Contribution and Feedback
Welcome to submit issue reports, feature requests, or contribute code! If you have any questions or suggestions, please contact us through GitHub Issues.

## License

This work is licensed under a [Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License](http://creativecommons.org/licenses/by-nc-sa/4.0/).

This means you are free to:
* Share — copy and redistribute the material in any medium or format
* Adapt — remix, transform, and build upon the material

Under the following terms:
* Attribution — You must give appropriate credit
* NonCommercial — You may not use the material for commercial purposes
* ShareAlike — If you remix, transform, or build upon the material, you must distribute your contributions under the same license
