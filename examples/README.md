# DOCX Table Converter 示例教程

本教程将指导您如何使用 DOCX Table Converter 包的各种功能。我们提供了多个示例，涵盖了从基本使用到高级特性的所有功能。

## 目录结构

```
examples/
├── README.md                      # 本教程文件
├── example.py                     # 示例代码
├── sample_data_single_header.csv  # 单行表头示例数据
├── sample_data_double_header.csv  # 双行表头示例数据
└── output/                        # 示例输出目录
```

## 示例数据说明

1. `sample_data_single_header.csv`
   - 单行表头的数据文件
   - 包含样品编号、温度、浓度、pH值和备注等字段
   - 适合基本的三线表转换

2. `sample_data_double_header.csv`
   - 双行表头的数据文件
   - 包含实验组和对照组的多个指标数据
   - 展示多级表头的处理能力

## 示例代码说明

`example.py` 文件包含了5个完整的示例，展示了不同的功能和使用场景：

### 示例1：基本功能
- 从CSV文件创建简单的三线表
- 使用单行表头
- 设置基本的字体和格式

```python
def example_1_basic():
    df = read_table_from_file('examples/sample_data_single_header.csv')
    write_table_to_docx(
        df=df,
        output_path='examples/output/example1.docx',
        table_caption='表 1',
        table_description='使用单行表头的基本三线表示例',
        font_name='宋体',
        font_size=10.5
    )
```

### 示例2：多级表头
- 处理双行表头数据
- 包含行索引
- 展示表头合并功能

```python
def example_2_multi_header():
    df = read_table_from_file(
        'examples/sample_data_double_header.csv',
        header_rows=2  # 指定双行表头
    )
    write_table_to_docx(
        df=df,
        output_path='examples/output/example2.docx',
        table_caption='表 2',
        table_description='使用双行表头的三线表示例',
        include_index=True  # 包含行索引
    )
```

### 示例3：特殊格式
- 处理上标、下标等特殊格式
- 使用不同的字体
- 展示科学计量单位的正确显示

```python
def example_3_special_formatting():
    df = pd.DataFrame({
        '温度': ['25 °C', '30 °C', '35 °C'],
        '浓度': ['100 mg·L⁻¹', '200 mg·L⁻¹', '300 mg·L⁻¹'],
        'pH': ['7.0±0.1', '7.5±0.2', '8.0±0.1'],
        'CO₂': ['5%', '10%', '15%'],
        'Ca²⁺': ['1.2', '1.5', '1.8']
    })
    write_table_to_docx(
        df=df,
        output_path='examples/output/example3.docx',
        special_formatting=True  # 启用特殊格式处理
    )
```

### 示例4：批处理
- 同时处理多个表格
- 支持合并为一个文件或分别保存
- 为每个表格设置不同的标题和描述

```python
def example_4_batch_processing():
    tables = [
        'examples/sample_data_single_header.csv',
        'examples/sample_data_double_header.csv'
    ]
    write_tables_to_docx(
        tables=tables,
        output_path='examples/output/example4_combined.docx',
        separate_files=False  # 合并为一个文件
    )
```

### 示例5：高级格式
- 自定义表格样式
- 设置页面布局
- 调整边框和背景

```python
def example_5_advanced_formatting():
    write_table_to_docx(
        df=df,
        output_path='examples/output/example5.docx',
        caption_bold=True,          # 标题加粗
        header_bold=True,           # 表头加粗
        table_width=6.0,            # 表格宽度
        header_background_color='F2F2F2',  # 表头背景色
        page_size='A4'             # 页面大小
    )
```

## 运行示例

1. 确保您已安装所有依赖：
   ```bash
   pip install -e .
   ```

2. 进入示例目录：
   ```bash
   cd examples
   ```

3. 运行示例代码：
   ```bash
   python example.py
   ```

4. 查看输出结果：
   - 所有生成的文件都在 `output` 目录中
   - 每个示例都会生成对应的 DOCX 文件
   - 批处理示例会同时生成合并文件和单独文件

## 注意事项

1. 特殊格式处理
   - 使用 `special_formatting=True` 启用特殊格式支持
   - 支持上标（⁺, ⁻）、下标（₁, ₂）等
   - 正确处理单位（°C, mg·L⁻¹）等

2. 表头设置
   - 单行表头：默认 `header_rows=1`
   - 双行表头：设置 `header_rows=2`
   - 无表头：设置 `header_rows=0`

3. 索引处理
   - 使用 `include_index=True` 包含行索引
   - 可以通过 GUI 选择特定列作为索引

4. 批处理注意事项
   - 支持混合处理 CSV 和 Excel 文件
   - 可以选择合并或分别保存
   - 每个表格可以有独立的标题和描述

## 常见问题

1. **表格宽度调整**
   - 使用 `table_width` 设置整体宽度
   - 可以通过 GUI 拖动调整列宽

2. **字体问题**
   - 确保系统安装了所需字体
   - 中文推荐使用"宋体"
   - 英文推荐使用"Times New Roman"

3. **特殊字符显示**
   - 确保使用正确的 Unicode 字符
   - 启用 `special_formatting` 功能

4. **文件保存**
   - 注意检查输出路径是否存在
   - 确保有写入权限
   - 建议使用绝对路径

## 更多信息

如需了解更多信息，请参考：
- [项目主页](https://github.com/JiehaoZhang1993/docx_table_converter)
- [完整文档](https://github.com/JiehaoZhang1993/docx_table_converter/wiki)
- [问题反馈](https://github.com/JiehaoZhang1993/docx_table_converter/issues)