"""
Tests for the core functionality of the DOCX Table Converter.
"""

import os
import pandas as pd
import pytest
import sys
import os

# Add the src directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from docx_table_converter.core import write_table_to_docx, read_table_from_file, parse_clipboard_data


@pytest.fixture
def sample_df():
    return pd.DataFrame({
        'Column 1': [1, 2, 3],
        'Column 2': ['A', 'B', 'C'],
        'Column 3': ['Value $_{1}$', 'Value $_{2}$', 'Value $_{3}$']
    })


@pytest.fixture
def temp_docx(tmp_path):
    return str(tmp_path / "test.docx")


def test_write_table_to_docx(sample_df, temp_docx):
    write_table_to_docx(
        df=sample_df,
        output_path=temp_docx,
        table_caption="Table 1",
        table_description="Test table",
        font_name="Times New Roman",
        font_size=12
    )
    assert os.path.exists(temp_docx)


def test_parse_clipboard_data():
    text = "Column 1\tColumn 2\tColumn 3\n1\tA\tValue $_{1}$\n2\tB\tValue $_{2}$"
    df = parse_clipboard_data(text)
    assert len(df) == 2
    assert len(df.columns) == 3


def test_read_table_from_file(sample_df, tmp_path):
    # Create a CSV file
    csv_path = tmp_path / "test.csv"
    sample_df.to_csv(csv_path, index=False)
    
    # Read the CSV file
    df = read_table_from_file(str(csv_path))
    assert len(df) == len(sample_df)
    assert len(df.columns) == len(sample_df.columns)


def test_special_formatting(sample_df, temp_docx):
    write_table_to_docx(
        df=sample_df,
        output_path=temp_docx,
        table_caption="Table 1",
        table_description="Test table with special formatting",
        special_formatting=True
    )
    assert os.path.exists(temp_docx)


def test_multiple_header_rows(temp_docx):
    # Create a DataFrame with multi-level columns
    arrays = [
        ['A', 'A', 'B', 'B'],
        ['1', '2', '1', '2']
    ]
    columns = pd.MultiIndex.from_arrays(arrays)
    df = pd.DataFrame(
        [[1, 2, 3, 4], [5, 6, 7, 8]],
        columns=columns
    )
    
    write_table_to_docx(
        df=df,
        output_path=temp_docx,
        table_caption="Table 1",
        table_description="Test table with multiple header rows"
    )
    assert os.path.exists(temp_docx)


def test_invalid_file_format():
    with pytest.raises(ValueError):
        read_table_from_file("invalid.txt") 

def main():
    # Run all tests
    pytest.main([__file__])