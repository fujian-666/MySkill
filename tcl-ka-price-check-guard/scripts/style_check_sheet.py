# -*- coding: utf-8 -*-
"""核对情况表「标准样式」固化器（以人工调好的定版格式为准）。

用法（COM 场景，核对情况表已生成后调用）：
    from style_check_sheet import apply_standard
    apply_standard(cs)          # cs = wb.Sheets('核对情况表')

规格（COM Interior.Color 原值，保证与人工调好的文件 100% 一致；不要用 openpyxl 重写颜色）：
  列宽   A=4.4 B=22 C=4.4 D=22 E=9.4 F=60 G=22 H=22
  标题   A1:H1 合并，黄底 65535，宋体11粗黑字，水平垂直居中，四边细边框
  表头   A2:H2，浅橙底 14277081，宋体11粗，居中，四边细边框
  条目行 白底，宋体11，垂直居中，四边细边框，行高 42.3，F列(异常说明)自动换行
  末两列 G/H 淡黄 13431551（业务填写区）
  结论行 A:H 合并，无边框，宋体12粗，绿✅/红❌，行高 15.3
"""
import re

FONT = '宋体'
COLW = {'A': 4.4, 'B': 22.0, 'C': 4.4, 'D': 22.0, 'E': 9.4, 'F': 60.0, 'G': 22.0, 'H': 22.0}
TITLE_FILL = 65535        # 黄 #FFFF00
HEADER_FILL = 14277081    # 浅橙（表头）
ITEM_FILL = 13431551      # 末两列（业务填写区）淡黄
XL_THIN, XL_LEFT, XL_RIGHT, XL_TOP, XL_BOTTOM = 1, 7, 8, 9, 10
XL_HCENTER, XL_VCENTER = -4108, -4107
XL_NONE = -4142


def _borders(rng, style=XL_THIN):
    for side in (XL_LEFT, XL_RIGHT, XL_TOP, XL_BOTTOM):
        rng.Borders(side).LineStyle = style


def _regions(cs):
    """探测：表头行=2；条目区=3..最后一个B列有值行；结论区=A列以✅/👉开头的行。
    用 UsedRange.Row + Rows.Count - 1 求末行，禁用 End(-4162)（AutoFilter 会漏行）。
    """
    last = cs.UsedRange.Rows.Count + cs.UsedRange.Row - 1
    items_end = 2
    res_rows = []
    for r in range(3, last + 1):
        b = cs.Cells(r, 2).Value
        a = str(cs.Cells(r, 1).Value or '')
        if b not in (None, ''):
            items_end = r
        if a.startswith(('✅', '👉')):
            res_rows.append(r)
    return items_end, res_rows, last


def apply_standard(cs):
    """对核对情况表应用标准样式（幂等，可重复调用）"""
    items_end, res_rows, last = _regions(cs)
    for col, w in COLW.items():
        cs.Columns(col).ColumnWidth = w
    t = cs.Range('A1:H1')
    t.Merge()
    t.Interior.Color = TITLE_FILL
    t.Font.Name = FONT
    t.Font.Size = 11
    t.Font.Bold = True
    t.HorizontalAlignment = XL_HCENTER
    t.VerticalAlignment = XL_VCENTER
    _borders(t)
    h = cs.Range('A2:H2')
    h.Interior.Color = HEADER_FILL
    h.Font.Name = FONT
    h.Font.Size = 11
    h.Font.Bold = True
    h.HorizontalAlignment = XL_HCENTER
    h.VerticalAlignment = XL_VCENTER
    _borders(h)
    if items_end >= 3:
        body = cs.Range('A3:H%d' % items_end)
        body.Font.Name = FONT
        body.Font.Size = 11
        body.VerticalAlignment = XL_VCENTER
        _borders(body)
        cs.Range('F3:F%d' % items_end).WrapText = True
        cs.Range('G3:H%d' % items_end).Interior.Color = ITEM_FILL
        for r in range(3, items_end + 1):
            if cs.Rows(r).RowHeight < 42.3 or cs.Rows(r).RowHeight > 60:
                cs.Rows(r).RowHeight = 42.3
    for r in res_rows:
        rng = cs.Range(cs.Cells(r, 1), cs.Cells(r, 8))
        if not rng.MergeCells:
            rng.Merge()
        for side in (XL_LEFT, XL_RIGHT, XL_TOP, XL_BOTTOM):
            rng.Borders(side).LineStyle = XL_NONE
        rng.Font.Name = FONT
        rng.Font.Size = 12
        rng.Font.Bold = True
        rng.VerticalAlignment = XL_VCENTER
        cs.Rows(r).RowHeight = 15.3
    return dict(items_end=items_end, res_rows=res_rows, last=last)


def restyle_sheet(cs):
    """别名：对外主入口"""
    return apply_standard(cs)


if __name__ == '__main__':
    import sys, os
    import win32com.client as wc
    f = os.path.abspath(sys.argv[1]).replace('/', '\\')
    excel = wc.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(f, 0)
        info = apply_standard(wb.Sheets('核对情况表'))
        wb.SaveAs(f, 51)
        wb.Close(True)
        print('styled:', info)
    finally:
        excel.Quit()
