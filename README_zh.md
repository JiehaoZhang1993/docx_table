# DOCX表格转换器

[English](README.md) | [中文](README_zh.md)

一个简单易用的工具，帮你把Excel、CSV等格式的表格转换成Word文档中的精美表格，再也不用为论文、报告中的表格格式发愁了！

## 为什么需要这个工具？

你是否曾经遇到过这些烦恼：
- 从Excel复制表格到Word后，格式全乱了
- 论文中的表格需要三线表格式，手动调整太麻烦
- 需要处理大量表格数据，一个个调整格式太耗时
- 表格中有特殊格式（如下标），在Word中难以正确显示
- 需要批量处理多个表格，但又想保持统一的格式

这个工具就是为了解决这些问题而生的！它能够：
- 自动将Excel、CSV表格转换为Word文档中的表格
- 支持三线表格格式（无竖线，只有横线）
- 自动处理特殊格式，如下标（如mg L$_{-1}$）
- 支持多级表头，适合复杂的数据展示
- 支持批量处理多个表格
- 提供友好的图形界面，无需编程知识也能使用

## 功能特点

- 支持从CSV、XLS、XLSX文件导入表格
- 支持直接从剪贴板粘贴表格数据
- 导出前可预览表格效果
- 自定义表格格式：
  - 中英文字体分别设置
  - 字体大小调整
  - 表格边框宽度可调
  - 表格标题可选是否加粗
- 支持多级表头（最多2级）
- 支持特殊格式（如下标）
- 三线表格式（无竖线）
- 双语界面（中文和英文）
- 批量处理功能：
  - 支持多个表格同时导入
  - 可选择合并导出或分别保存
  - 每个表格可独立设置标题和描述
  - 支持预览所有表格效果
- 保存功能：
  - 支持在预览界面直接保存
  - 可选择保存位置
  - 自动打开生成的文档

## 安装方法

目前项目尚未发布到PyPI，你可以通过以下方式安装：

### 方法一：从GitHub克隆

```bash
git clone https://github.com/JiehaoZhang1993/docx_table.git
cd docx_table
pip install -e .
```

### 方法二：直接下载源码

1. 下载并解压项目源码
2. 进入项目目录
3. 运行 `pip install -e .` 安装

## 使用方法

### 方法一：作为Python包使用

如果你熟悉Python，可以直接在代码中调用：

```python
import pandas as pd
from docx_table_converter import write_table_to_docx, write_tables_to_docx

# 单个表格处理
df = pd.DataFrame({
    '样本': ['A', 'B', 'C'],
    '温度 (°C)': [25, 30, 35],
    '浓度 (mg L$_{-1}$)': [10, 20, 30],
    'pH值': [7.0, 7.2, 7.4]
})

# 写入单个表格
write_table_to_docx(
    df=df, 
    output_path='输出表格.docx',
    table_caption='表1. 样本数据',
    chinese_font='宋体',
    english_font='Times New Roman',
    font_size=12,
    border_width=1,
    bold_caption=True,
    mode='append'
)

# 批量处理多个表格
tables = [df1, df2, df3]  # DataFrame列表或文件路径列表
captions = ['表1', '表2', '表3']
write_tables_to_docx(
    tables=tables,
    captions=captions,
    output_path='批量表格.docx',
    font_name='Times New Roman',
    font_size=10.5,
    header_rows=[1, 1, 1],  # 每个表格的表头行数
    separate_files=False  # True则每个表格保存为单独文件
)
```

### 方法二：使用图形界面

如果你不熟悉编程，可以使用图形界面：

1. 运行程序：`python run_gui.py`
2. 在主界面中：
   - 点击"浏览"选择Excel或CSV文件，或点击"粘贴表格数据"从剪贴板导入
   - 设置表头行数（0-2行）
   - 自定义中英文字体、字号
   - 设置表格标题
   - 调整表格边框宽度
   - 选择是否加粗标题
3. 点击"预览"查看效果，可以在预览界面直接保存
4. 点击"导出"保存为Word文档

批量处理：
1. 点击"批量处理"按钮
2. 在批量处理界面：
   - 添加多个文件或粘贴多个表格
   - 为每个表格设置标题和描述
   - 选择是否合并保存
3. 可以预览单个表格或所有表格的效果
4. 点击"导出"完成批量转换

## 参数说明

### write_table_to_docx 函数参数

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| df | DataFrame | 必填 | 要转换的pandas DataFrame |
| output_path | str | 必填 | 输出DOCX文件路径 |
| table_caption | str | "Table 1" | 表格标题 |
| chinese_font | str | '宋体' | 中文字体 |
| english_font | str | 'Times New Roman' | 英文字体 |
| font_size | int | 12 | 字号 |
| header_rows | int | 1 | 表头行数 |
| mode | str | 'append' | 文件模式：'append'（追加）或'overwrite'（覆盖） |
| border_width | float | 1 | 表格边框宽度 |
| bold_caption | bool | True | 是否加粗表格标题 |

### write_tables_to_docx 函数参数

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| tables | list | 必填 | DataFrame列表或文件路径列表 |
| captions | list | 必填 | 表格标题列表 |
| output_path | str | 必填 | 输出文件路径 |
| font_name | str | 'Times New Roman' | 字体名称 |
| font_size | float | 10.5 | 字号 |
| header_rows | list | None | 每个表格的表头行数列表 |
| separate_files | bool | False | 是否将每个表格保存为单独的文件 |

## 贡献与反馈

欢迎提交问题报告、功能请求或贡献代码！如果你有任何问题或建议，请通过GitHub Issues与我们联系。

## 许可证

 CC BY-NC-SA 4.0 International License  
[![CC BY-NC-SA 4.0](https://licensebuttons.net/l/by-nc-sa/4.0/88x31.png)](https://creativecommons.org/licenses/by-nc-sa/4.0/)