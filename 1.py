from docx import Document
from docx.shared import Inches, Pt
import pandas as pd
from docx.enum.table import WD_TABLE_ALIGNMENT
import numpy as np


def merge_cells_by_column(table, column_index):
    cells = [row.cells[column_index] for row in table.rows]

    for i in range(len(cells) - 1):
        if cells[i].text == cells[i + 1].text:
            a = cells[i + 1].merge(cells[i])


def generate_table(file_name, doc):
    df = pd.read_csv('data/{}.csv'.format(file_name))
    print('行数' + str(len(df)))
    print('列数' + str(len(df.columns)))
    # doc = Document(doc)
    table = doc.add_table(1, cols=len(df.columns), style='Table Grid')
    table.style.paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style.font.size = Pt(8)
    # table = table[0]
    header_cells = table.rows[0].cells
    for i, col in enumerate(df.columns):
        header_cells[i].text = col

    # 填入数据
    for i, row in enumerate(df.itertuples(), start=1):
        cells = table.add_row().cells
        for j, value in enumerate(row[1:], start=0):
            cells[j].text = str(value)

    # cell1 = table.cell(0, 0)
    # cell2 = table.cell(0, 1)
    # cell1.merge(cell2)
    merge_cells_by_column(table, 0)
    # merge_cells(table, 1, 0, len(df), 1)
    # merge_cells(table, 1, 0, 0, 1)

    # for row in table.rows:
    #     for cell in row.cells:
    #         cell.text = df[row][cell]
    # table = doc.add_table(rows=len(df), cols=len(df.columns))

if __name__ == '__main__':
    document = Document()
    document.add_heading('总结概述', level=0)
    document.add_heading('1. 整体风险', level=1)
    # print(type(document))
    # p1 = document.add_paragraph('整体风险', style='List Number')
    p1 = document.add_paragraph('本次报告分析周期为：（2020年01月01日到2023年02月28日），通过发票、财务报表、纳税申报表的综合分析，共检测出风险点14项')
    p1.paragraph_format.line_spacing = Pt(20)
    # p2 = document.add_paragraph('本次报告分析周期为：（2020年01月01日到2023年02月28日），通过发票、财务报表、纳税申报表的综合分析，共检测出风险点14项')
    # p1.paragraph_format.line_spacing = Pt(20)
    # p1.add_run('本次报告分析周期为：（2020年01月01日到2023年02月28日），通过发票、财务报表、纳税申报表的综合分析，共检测出风险点14项')
    # r1.font.size = Pt(10)
    document.add_heading('2. 具体风险如下', level=1)
    # table = document.add_table(rows=2, cols=2)
    # cell = table.cell(0, 1)
    # cell.text = 'parrot, possibly dead'
    # row = table.rows[1]
    # row.cells[0].text = 'Foo bar to you.'
    # row.cells[1].text = 'And a hearty foo bar to you too sir!'
    generate_table('1', document)


    # document.add_page_break()
    # p2 = document.add_paragraph('具体风险如下', style='List Number')

    document.save('new.docx')
    