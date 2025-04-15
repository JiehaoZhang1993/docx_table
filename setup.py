#!/usr/bin/env python
"""
Setup script for the DOCX Table Converter package.
"""

from setuptools import setup, find_packages
import os

# Read the README file if it exists
readme_path = os.path.join(os.path.dirname(__file__), "README.md")
if os.path.exists(readme_path):
    with open(readme_path, encoding="utf-8") as f:
        long_description = f.read()
else:
    long_description = "DOCX Table Converter - A tool for converting tables to DOCX format"

setup(
    name="docx-table-converter",
    version="0.1.0",
    description="A tool for converting tables to DOCX format with customizable formatting",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="JiehaoZhang",
    author_email="zhangjiehao1993@gmail.com",
    url="https://github.com/JiehaoZhang1993/docx_table_converter",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        "pandas>=1.0.0",
        "python-docx>=0.8.11",
        "openpyxl>=3.0.0",
        "xlrd>=2.0.0",
        "PyQt5>=5.15.0",
    ],
    extras_require={
        "dev": ["pytest", "black", "isort", "flake8"],
    },
    entry_points={
        "console_scripts": [
            "docx-table-converter=docx_table_converter.gui:main",
        ],
    },
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Science/Research",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Topic :: Office/Business :: Word Processors",
    ],
    python_requires=">=3.6",
) 