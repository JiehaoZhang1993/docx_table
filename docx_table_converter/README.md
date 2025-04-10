# DOCX Table Converter

A tool to convert tables from various formats (CSV, XLS, XLSX) to research paper format with three-line tables in DOCX.

## Features

- Import tables from CSV, XLS, and XLSX files
- Convert tables to research paper format with three-line tables
- Customize table captions and numbering
- Support for special formatting (subscripts, superscripts)
- GUI application for easy use
- Python package for advanced users

## Installation

### As a Python Package

```bash
pip install docx-table-converter
```

### From Source

```bash
git clone https://github.com/yourusername/docx_table_converter.git
cd docx_table_converter
pip install -e .
```

## Usage

### GUI Application

Run the GUI application:

```bash
docx-table-converter
```

Or:

```bash
python -m docx_table_converter.gui
```

### Python Package

```python
import pandas as pd
from docx_table_converter import write_table_to_docx

# Create a DataFrame
df = pd.DataFrame({
    'Column 1': [1, 2, 3],
    'Column 2': ['A', 'B', 'C'],
    'Column 3': ['Value $_{1}$', 'Value $_{2}$', 'Value $_{3}$']
})

# Convert to DOCX with three-line table
write_table_to_docx(
    df=df,
    output_path='output.docx',
    table_caption='Table 1',
    table_description='Example table with special formatting',
    font_name='Times New Roman',
    font_size=12
)
```

## Advanced Usage

The package provides many customization options for advanced users:

```python
from docx_table_converter import write_table_to_docx

write_table_to_docx(
    df=df,
    output_path='output.docx',
    table_caption='Table 1',
    table_description='Example table with special formatting',
    headers=['Custom Header 1', 'Custom Header 2', 'Custom Header 3'],
    font_name='Times New Roman',
    font_size=12,
    include_index=True,
    mode='append',
    caption_bold=True,
    header_bold=True,
    header_rows=2,
    table_width=6.0,
    column_widths=[2.0, 2.0, 2.0],
    cell_padding=0.1,
    border_width=1.0,
    border_color='000000',
    header_background_color='F2F2F2',
    caption_alignment='left',
    table_alignment='center',
    page_margins=(2.5, 2.5, 2.5, 2.5),
    page_size='A4',
    page_orientation='portrait',
    page_break_after=True,
    special_formatting=True,
    subscript_pattern=r'\$_{(.+?)}\$',
    superscript_pattern=r'\$^{(.+?)}\$',
    italic_pattern=r'\*_(.+?)_\*',
    bold_pattern=r'\*\*(.+?)\*\*',
    underline_pattern=r'__(.+?)__',
    strikethrough_pattern=r'~~(.+?)~~',
    highlight_pattern=r'==(.+?)==',
    color_pattern=r'\{color:(.+?)\}(.+?)\{/color\}',
    font_pattern=r'\{font:(.+?)\}(.+?)\{/font\}',
    size_pattern=r'\{size:(.+?)\}(.+?)\{/size\}',
    custom_patterns=None,
    custom_formatters=None,
    custom_styles=None,
    custom_borders=None,
    custom_captions=None,
    custom_headers=None,
    custom_data=None,
    custom_document=None,
    custom_table=None,
    custom_cell=None,
    custom_paragraph=None,
    custom_run=None,
    custom_style=None,
    custom_font=None,
    custom_border=None,
    custom_caption=None,
    custom_header=None,
    custom_data_cell=None,
    custom_document_settings=None,
    custom_table_settings=None,
    custom_cell_settings=None,
    custom_paragraph_settings=None,
    custom_run_settings=None,
    custom_style_settings=None,
    custom_font_settings=None,
    custom_border_settings=None,
    custom_caption_settings=None,
    custom_header_settings=None,
    custom_data_cell_settings=None,
)
```

## License

This project is licensed under the MIT License - see the LICENSE file for details. 