import os
import pytest
from openpyxl import Workbook
from tools import _generate_sheets_markdown_summary


# 文件名：test_example.py
def test_generate_sheets_markdown_summary():
    # 创建一个临时 Excel 文件
    markdown_str = _generate_sheets_markdown_summary('example_table.xlsx')
    print(markdown_str)
    # 验证输出
    assert "📊 **Excel File Overview:" in markdown_str
    assert "**Total Sheets:** 1" in markdown_str
