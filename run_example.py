#!/usr/bin/env python
"""
Script to run the DOCX Table Converter example.
"""

import sys
import os

# Add the src directory to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), 'src')))

from examples.example import main

if __name__ == '__main__':
    main() 