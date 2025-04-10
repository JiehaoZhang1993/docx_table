"""
Core functionality for converting tables to DOCX format.
"""

import os
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import re
import io


def read_table_from_file(file_path, sheet_name=None, header_rows=1):
    """
    Read a table from a file (CSV, XLS, XLSX) and return it as a pandas DataFrame.
    
    Parameters:
    -----------
    file_path : str
        Path to the input file
    sheet_name : str or int, optional
        Name or index of the sheet to read (for Excel files). If None, reads the first sheet.
    header_rows : int, optional
        Number of header rows in the table (default: 1)
    
    Returns:
    --------
    pandas.DataFrame
        The table data as a DataFrame
    """
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext == '.csv':
        return pd.read_csv(file_path, header=list(range(header_rows)))
    elif file_ext in ['.xls', '.xlsx']:
        # If sheet_name is None, read the first sheet
        if sheet_name is None:
            return pd.read_excel(file_path, header=list(range(header_rows)))
        return pd.read_excel(file_path, sheet_name=sheet_name, header=list(range(header_rows)))
    else:
        raise ValueError(f"Unsupported file format: {file_ext}")


def parse_clipboard_data(text, header_rows=1):
    """
    Parse table data from clipboard text.
    
    Parameters:
    -----------
    text : str
        Table data as text (tab or comma separated)
    header_rows : int, optional
        Number of header rows (default: 1)
    
    Returns:
    --------
    pandas.DataFrame
        The table data as a DataFrame
    """
    # Try to detect the delimiter
    sample_line = text.split('\n')[0]
    if '\t' in sample_line:
        delimiter = '\t'
    else:
        delimiter = ','
    
    # Convert text to DataFrame
    buffer = io.StringIO(text)
    df = pd.read_csv(buffer, sep=delimiter, header=list(range(header_rows-1)))
    
    return df


def write_table_to_docx(df, output_path, table_caption="Table", table_description="", 
                        headers=None, font_name='Times New Roman', font_size=12, include_index=False,
                        mode='append'):
    """
    Write a pandas DataFrame as a formatted table to a Word document.
    If the document exists, it will append the table to it or overwrite it based on mode parameter.
    
    Parameters:
    -----------
    df : pandas.DataFrame
        The DataFrame to be written as a table
    output_path : str
        Path where the Word document should be saved
    table_caption : str, optional
        Caption number/identifier for the table (default: "Table")
    table_description : str, optional
        Description text for the table (default: "")
    headers : list, optional
        Custom headers for the table. If None, DataFrame column names will be used
    font_name : str, optional
        Font to use in the document (default: 'Times New Roman')
    font_size : int, optional
        Font size in points (default: 12)
    include_index : bool, optional
        Whether to include the DataFrame index in the table (default: False)
    mode : str, optional
        Mode to use when file exists: 'append' (add to existing document) or 'overwrite' (create new document)
        (default: 'append')
    
    Returns:
    --------
    None
    """
    # Check if document exists
    if os.path.exists(output_path) and mode == 'append':
        doc = Document(output_path)
    else:
        # Create a new document
        doc = Document()
        
        # Set page margins (2.5 cm)
        sections = doc.sections
        for section in sections:
            section.left_margin = Cm(2.5)
            section.right_margin = Cm(2.5)
        
        # Set default font
        style = doc.styles['Normal']
        font = style.font
        font.name = font_name
        if hasattr(style, '_element') and hasattr(style._element, 'rPr'):
            style._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
        font.size = Pt(font_size)
    
    # Add table caption
    caption = doc.add_paragraph()
    caption.alignment = WD_ALIGN_PARAGRAPH.LEFT
    bold_run = caption.add_run(f'{table_caption}. ')
    bold_run.bold = True
    bold_run.font.name = font_name
    bold_run.font.size = Pt(font_size)
    
    if table_description:
        normal_run = caption.add_run(table_description)
        normal_run.bold = False
        normal_run.font.name = font_name
        normal_run.font.size = Pt(font_size)
    
    # Check if DataFrame has MultiIndex columns
    has_multi_level_columns = isinstance(df.columns, pd.MultiIndex)
    
    # Initialize num_header_rows to avoid UnboundLocalError
    num_header_rows = 1
    
    # Prepare data for table
    if include_index:
        # Create a copy of the DataFrame with index as a column
        df_with_index = df.copy()
        df_with_index = df_with_index.reset_index()
        table_cols = len(df_with_index.columns)
        
        # Set headers for table with index
        if headers is None:
            headers = df_with_index.columns
    else:
        # Use original DataFrame without index
        table_cols = len(df.columns)
        
        # Set headers for table without index
        if headers is None:
            headers = df.columns
    
    data_df = df_with_index if include_index else df
    
    # Handle multi-level columns
    if has_multi_level_columns:
        num_header_rows = len(data_df.columns.levels)
        rows = min(len(data_df) + num_header_rows, 1000)  # Limit rows to prevent index errors
        table = doc.add_table(rows=rows, cols=table_cols)
        
        # Add multi-level headers
        for level in range(num_header_rows):
            # Track cells to merge for each level
            merge_ranges = []
            start_idx = 0
            current_val = None
            
            # Identify ranges to merge
            for col_idx, col_val in enumerate(data_df.columns.get_level_values(level)):
                if col_val != current_val:
                    if col_idx > start_idx and current_val is not None:
                        merge_ranges.append((start_idx, col_idx - 1))
                    current_val = col_val
                    start_idx = col_idx
                
                # Add the text to the cell
                cell = table.cell(level, col_idx)
                cell.text = str(col_val)
                paragraph = cell.paragraphs[0]
                run = paragraph.runs[0]
                run.font.bold = True
                run.font.name = font_name
                run.font.size = Pt(font_size)
            
            # Add the last range if needed
            if start_idx < table_cols - 1 and current_val is not None:
                merge_ranges.append((start_idx, table_cols - 1))
            
            # Perform the merges
            for start, end in merge_ranges:
                if start != end:  # Only merge if there's more than one cell
                    table.cell(level, start).merge(table.cell(level, end))
        
        # Adjust data row start index
        data_row_start = num_header_rows

    else:
        rows = min(len(data_df) + 1, 1000)  # Limit rows to prevent index errors
        table = doc.add_table(rows=rows, cols=table_cols)
        
        # Add single-level headers
        for j, header in enumerate(headers):
            cell = table.cell(0, j)
            cell.text = str(header)
            paragraph = cell.paragraphs[0]
            run = paragraph.runs[0]
            run.font.bold = True
            run.font.name = font_name
            run.font.size = Pt(font_size)
        
        # Adjust data row start index
        data_row_start = 1
    
    table.style = 'Table Grid'
    
    # Handle table style safely
    if hasattr(table, 'style') and hasattr(table.style, 'element'):
        borders = table.style.element.xpath('//w:tblBorders')
        if borders and len(borders) > 0:
            parent = borders[0].getparent()
            if parent is not None:
                parent.remove(borders[0])
    
    # Add data
    for i, (idx, row) in enumerate(data_df.iterrows()):
        if i >= rows - data_row_start:  # Check if we're exceeding table dimensions
            break
        for j, value in enumerate(row):
            cell = table.cell(i + data_row_start, j)
            
            # Check if the value contains formatting like subscripts or superscripts
            value_str = str(value)
            
            # Clear existing text
            cell.text = ""
            paragraph = cell.paragraphs[0]
            
            # Handle subscripts (e.g., $_{text}$)
            sub_pattern = re.compile(r'\$_{(.+?)}\$')
            parts = sub_pattern.split(value_str)
            
            for k, part in enumerate(parts):
                if k % 2 == 0:  # Regular text
                    if part:
                        run = paragraph.add_run(part)
                        run.font.name = font_name
                        run.font.size = Pt(font_size)
                else:  # Subscript text
                    run = paragraph.add_run(part)
                    run.font.name = font_name
                    run.font.size = Pt(font_size)
                    run.font.subscript = True
            
            # If no formatting was found, add the text normally
            if len(parts) == 1:
                cell.text = value_str
                paragraph = cell.paragraphs[0]
                if paragraph.runs:
                    run = paragraph.runs[0]
                    run.font.name = font_name
                    run.font.size = Pt(font_size)
    
    # Apply three-line table style with no vertical lines
    for row_idx, row in enumerate(table.rows):
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            
            # Determine border style based on row position and multi-level headers
            top_border = 'none'
            bottom_border = 'none'
            
            if has_multi_level_columns:
                if row_idx == 0:  # First header row
                    top_border = 'single'
                elif row_idx == num_header_rows - 1:  # Last header row
                    bottom_border = 'single'
                elif row_idx == len(table.rows) - 1:  # Last data row
                    bottom_border = 'single'
            else:
                if row_idx == 0:  # Header row
                    top_border = 'single'
                    bottom_border = 'single'
                elif row_idx == len(table.rows) - 1:  # Last data row
                    bottom_border = 'single'
            
            borders_xml = f'''<w:tcBorders {nsdecls("w")}>
                <w:top w:val="{top_border}" w:sz="8"/>
                <w:bottom w:val="{bottom_border}" w:sz="8"/>
                <w:left w:val="none"/>
                <w:right w:val="none"/>
                <w:insideH w:val="none"/>
                <w:insideV w:val="none"/>
            </w:tcBorders>'''
            existing_borders = tcPr.xpath('./w:tcBorders')
            if existing_borders:
                tcPr.remove(existing_borders[0])
            tcBorders = parse_xml(borders_xml)
            tcPr.append(tcBorders)
    
    # Add page break after the table
    doc.add_paragraph().add_run().add_break()
    
    # Save document
    doc.save(output_path)
    # Open saved file with word or default application
    if os.name == 'nt':
        os.startfile(output_path)
    elif os.name == 'posix':
        os.system(f'xdg-open "{output_path}"')
    else:
        print(f"File saved at {output_path}. Please open it manually.")
    return None 