"""
Example script demonstrating how to use the DOCX Table Converter package.
"""

import pandas as pd
import sys
import os

# Add the src directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from docx_table_converter import write_table_to_docx, read_table_from_file


def main():
    # Example 1: Create a simple table with special formatting
    df = pd.DataFrame({
        'Sample': ['Sample 1', 'Sample 2', 'Sample 3'],
        'Temperature (°C)': [25, 30, 35],
        'Concentration (mg L$_{-1}$)': [100, 200, 300],
        'pH': [7.0, 7.5, 8.0]
    })
    
    write_table_to_docx(
        df=df,
        output_path='example1.docx',
        table_caption='Table 1',
        table_description='Sample data with temperature, concentration, and pH measurements.',
        font_name='Times New Roman',
        font_size=12,
        special_formatting=True
    )
    
    # Example 2: Create a table with multiple header rows
    arrays = [
        ['Control', 'Control', 'Treatment', 'Treatment'],
        ['Before', 'After', 'Before', 'After']
    ]
    columns = pd.MultiIndex.from_arrays(arrays)
    df_multi = pd.DataFrame(
        [
            [10.5, 11.2, 10.8, 12.5],
            [9.8, 10.5, 10.2, 13.1],
            [10.1, 10.8, 10.5, 12.8]
        ],
        columns=columns,
        index=['Sample 1', 'Sample 2', 'Sample 3']
    )
    
    write_table_to_docx(
        df=df_multi,
        output_path='example2.docx',
        table_caption='Table 2',
        table_description='Comparison of control and treatment groups before and after intervention.',
        font_name='Times New Roman',
        font_size=12,
        include_index=True
    )
    
    # # Example 3: Read from CSV and customize formatting
    # df_csv = read_table_from_file('data.csv')  # You need to create this CSV file
    
    # write_table_to_docx(
    #     df=df_csv,
    #     output_path='example3.docx',
    #     table_caption='Table 3',
    #     table_description='Data imported from CSV file with custom formatting.',
    #     font_name='Times New Roman',
    #     font_size=12,
    #     caption_bold=True,
    #     header_bold=True,
    #     table_width=6.0,
    #     column_widths=[1.5, 1.5, 1.5, 1.5],
    #     cell_padding=0.1,
    #     border_width=1.0,
    #     border_color='000000',
    #     header_background_color='F2F2F2',
    #     caption_alignment='left',
    #     table_alignment='center',
    #     page_margins=(2.5, 2.5, 2.5, 2.5),
    #     page_size='A4',
    #     page_orientation='portrait',
    #     page_break_after=True,
    #     special_formatting=True
    # )



if __name__ == '__main__':
    main() 