#!/usr/bin/env python
"""
启动DOCX表格转换器的GUI应用程序
"""

import sys
import os

# 添加src目录到Python路径
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), 'src')))

# 导入并运行GUI
from docx_table_converter.gui import main

if __name__ == "__main__":
    main() 