# DOCX Table Converter

A tool for converting tables from various formats (CSV, Excel) to DOCX format with customizable formatting.

## Features

- Import tables from CSV, XLS, and XLSX files
- Paste table data directly from clipboard
- Preview table data before export
- Customize table formatting (font, size, alignment, etc.)
- Support for multi-level headers
- Special formatting for subscripts (e.g., mg L$_{-1}$)
- Tables with no vertical lines (three-line table style)
- Bilingual interface (English and Chinese)

## Installation

```bash
pip install docx-table-converter
```

## Usage

### As a Python Package

```python
import pandas as pd
from docx_table_converter import write_table_to_docx

# Create a sample DataFrame
df = pd.DataFrame({
    'Sample': ['A', 'B', 'C'],
    'Temperature (°C)': [25, 30, 35],
    'Concentration (mg L$_{-1}$)': [10, 20, 30],
    'pH': [7.0, 7.2, 7.4]
})

# Write to DOCX
write_table_to_docx(
    df, 
    'output.docx',
    table_caption='Table 1. Sample data',
    table_description='Experimental results for different samples.',
    font_name='Times New Roman',
    font_size=12,
    header_bold=True,
    special_formatting=True
)
```

### As an Executable

Run the executable and use the GUI to:
1. Import a table from a file or paste from clipboard
2. Preview the table data
3. Customize formatting options
4. Export to DOCX

## License

MIT 