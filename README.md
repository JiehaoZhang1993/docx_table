# DOCX表格转换器

一个简单易用的工具，帮你把Excel、CSV等格式的表格转换成Word文档中的精美表格，再也不用为论文、报告中的表格格式发愁了！

## 为什么需要这个工具？

你是否曾经遇到过这些烦恼：
- 从Excel复制表格到Word后，格式全乱了
- 论文中的表格需要三线表格式，手动调整太麻烦
- 需要处理大量表格数据，一个个调整格式太耗时
- 表格中有特殊格式（如下标），在Word中难以正确显示

这个工具就是为了解决这些问题而生的！它能够：
- 自动将Excel、CSV表格转换为Word文档中的表格
- 支持三线表格式（无竖线，只有横线）
- 自动处理特殊格式，如下标（如mg L$_{-1}$）
- 支持多级表头，适合复杂的数据展示
- 提供友好的图形界面，无需编程知识也能使用

## 功能特点

- 支持从CSV、XLS、XLSX文件导入表格
- 支持直接从剪贴板粘贴表格数据
- 导出前可预览表格效果
- 自定义表格格式（字体、大小、对齐方式等）
- 支持多级表头
- 支持特殊格式（如下标）
- 三线表格式（无竖线）
- 双语界面（中文和英文）

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

### 方法一：作为Python包使用（适合程序员）

如果你熟悉Python，可以直接在代码中调用：

```python
import pandas as pd
from docx_table_converter import write_table_to_docx

# 创建一个示例DataFrame
df = pd.DataFrame({
    '样本': ['A', 'B', 'C'],
    '温度 (°C)': [25, 30, 35],
    '浓度 (mg L$_{-1}$)': [10, 20, 30],
    'pH值': [7.0, 7.2, 7.4]
})

# 写入DOCX文件
write_table_to_docx(
    df, 
    '输出表格.docx',
    table_caption='表1. 样本数据',
    table_description='不同样本的实验结果。',
    font_name='宋体',  # 支持中文字体
    font_size=12,
    include_index=False,  # 是否包含索引列
    mode='append'  # 追加到现有文档或创建新文档
)
```

### 方法二：使用图形界面（适合所有用户）

如果你不熟悉编程，可以使用图形界面：

1. 运行程序：`python run_gui.py`
2. 在界面中：
   - 点击"浏览"选择Excel或CSV文件，或点击"粘贴表格数据"从剪贴板导入
   - 设置表头行数
   - 自定义字体、字号、表格标题和描述
   - 选择是否包含索引列
   - 选择导出模式（追加或覆盖）
3. 点击"预览"查看效果
4. 点击"导出"保存为Word文档

<img src="./img/1744277362148.png" width="400" alt="GUI界面预览">

## 参数说明

`write_table_to_docx` 函数支持以下参数：

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| df | DataFrame | 必填 | 要转换的pandas DataFrame |
| output_path | str | 必填 | 输出DOCX文件路径 |
| table_caption | str | "Table" | 表格标题 |
| table_description | str | "" | 表格描述文本 |
| headers | list | None | 自定义表头，默认使用DataFrame的列名 |
| font_name | str | 'Times New Roman' | 字体名称 |
| font_size | int | 12 | 字号 |
| include_index | bool | False | 是否包含DataFrame的索引列 |
| mode | str | 'append' | 文件模式：'append'（追加）或'overwrite'（覆盖） |

## 贡献与反馈

欢迎提交问题报告、功能请求或贡献代码！如果你有任何问题或建议，请通过GitHub Issues与我们联系。

## 许可证

MIT
