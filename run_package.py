#!/usr/bin/env python
"""
Script to run the DOCX Table Converter package directly.
"""

import sys
import os

# Add the src directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), 'src')))

from docx_table_converter.gui import main

if __name__ == '__main__':
    main() 