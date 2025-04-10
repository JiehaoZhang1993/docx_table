from setuptools import setup, find_packages

setup(
    name="docx_table_converter",
    version="0.1.0",
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
    author="Jiehao Zhang",
    author_email="zhangjiehao1993@gmail.com",
    description="A tool to convert tables to research paper format with three-line tables in DOCX",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/docx_table_converter",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",
) 