"""
示例代码：展示如何使用 DOCX Table Converter 包的各种功能。
"""

import pandas as pd
import sys
import os

# 将 src 目录添加到 Python 路径
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../src')))

from docx_table_converter import write_table_to_docx, read_table_from_file, write_tables_to_docx


def example_1_basic():
    """基本功能示例：从CSV文件创建简单的三线表"""
    # 创建输出目录
    os.makedirs(os.path.abspath(os.path.join(os.path.dirname(__file__), 'output')), exist_ok=True)
    
    # 读取单行表头的CSV文件
    df = read_table_from_file(os.path.abspath(os.path.join(os.path.dirname(__file__), 'sample_data_single_header.csv')))
    
    # 导出为DOCX文件
    write_table_to_docx(
        df=df,
        output_path=os.path.abspath(os.path.join(os.path.dirname(__file__), 'output/example1.docx')),
        table_caption='表 1',
        table_description='使用单行表头的基本三线表示例',
        font_name='宋体',
        font_size=10.5,
        include_index=False
    )


def example_2_multi_header():
    """多级表头示例：处理双行表头数据"""
    # 创建输出目录
    os.makedirs(os.path.abspath(os.path.join(os.path.dirname(__file__), 'output')), exist_ok=True)
    
    # 读取双行表头的CSV文件
    df = read_table_from_file(
        os.path.abspath(os.path.join(os.path.dirname(__file__), 'sample_data_double_header.csv')),
        header_rows=2  # 指定双行表头
    )
    
    # 导出为DOCX文件
    write_table_to_docx(
        df=df,
        output_path=os.path.abspath(os.path.join(os.path.dirname(__file__), 'output/example2.docx')),
        table_caption='表 2',
        table_description='使用双行表头的三线表示例',
        font_name='宋体',
        font_size=10.5,
        include_index=True  # 包含行索引
    )


def example_3_special_formatting():
    """特殊格式示例：包含上标、下标等特殊格式"""
    # 创建输出目录
    os.makedirs(os.path.abspath(os.path.join(os.path.dirname(__file__), 'output')), exist_ok=True)
    
    # 创建包含特殊格式的DataFrame
    df = pd.DataFrame({
        '温度': ['25 °C', '30 °C', '35 °C'],
        '浓度': ['100 mg·L⁻¹', '200 mg·L⁻¹', '300 mg·L⁻¹'],
        'pH': ['7.0±0.1', '7.5±0.2', '8.0±0.1'],
        'CO₂': ['5%', '10%', '15%'],
        'Ca²⁺': ['1.2', '1.5', '1.8']
    })
    
    # 导出为DOCX文件
    write_table_to_docx(
        df=df,
        output_path=os.path.abspath(os.path.join(os.path.dirname(__file__), 'output/example3.docx')),
        table_caption='表 3',
        table_description='包含特殊格式（上标、下标等）的三线表示例',
        font_name='Times New Roman',  # 使用英文字体
        font_size=10.5
        # 移除了special_formatting参数，因为它不被支持
    )


def example_4_batch_processing():
    """批处理示例：同时处理多个表格"""
    # 创建输出目录
    os.makedirs(os.path.abspath(os.path.join(os.path.dirname(__file__), 'output')), exist_ok=True)
    os.makedirs(os.path.abspath(os.path.join(os.path.dirname(__file__), 'output/separate')), exist_ok=True)
    
    # 准备多个表格数据
    tables = [
        os.path.abspath(os.path.join(os.path.dirname(__file__), 'sample_data_single_header.csv')),
        os.path.abspath(os.path.join(os.path.dirname(__file__), 'sample_data_double_header.csv'))
    ]
    
    # 设置表格标题
    captions =['单行表头数据',
         '双行表头数据']
    
    # 设置表格描述
    descriptions =        ['这是第一个表格的详细描述',        '这是第二个表格的详细描述']
    
    header_rows=[1,2]
    
    # 批量导出 - 合并为一个文件
    write_tables_to_docx(
        tables=tables,
        captions=captions,
        output_path=os.path.abspath(os.path.join(os.path.dirname(__file__), 'output/example4_combined.docx')),
        descriptions=descriptions,
        font_name='宋体',
        font_size=10.5,
        separate_files=False,  # 合并为一个文件
        header_rows=header_rows  # 指定表头行数
    )
    
    # 批量导出 - 每个表格单独保存
    write_tables_to_docx(
        tables=tables,
        captions=captions,
        output_path=os.path.abspath(os.path.join(os.path.dirname(__file__), 'output/separate')),  # 指定输出目录
        descriptions=descriptions,
        font_name='宋体',
        font_size=10.5,
        separate_files=True,  # 分别保存
        header_rows=header_rows
    )




def main():
    """运行所有示例"""
    # 创建输出目录
    
    os.makedirs(os.path.abspath(os.path.join(os.path.dirname(__file__), 'output')), exist_ok=True)
    os.makedirs(os.path.abspath(os.path.join(os.path.dirname(__file__), 'output/separate')), exist_ok=True)
    
    # 运行示例
    print("运行示例 1: 基本功能...")
    example_1_basic()
    
    print("运行示例 2: 多级表头...")
    example_2_multi_header()
    
    print("运行示例 3: 特殊格式...")
    example_3_special_formatting()
    
    print("运行示例 4: 批处理...")
    example_4_batch_processing()

    print("\n所有示例已完成！请查看 output 目录下的结果。")


if __name__ == '__main__':
    main()