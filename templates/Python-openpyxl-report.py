# Python-openpyxl-report.py
# 对应教程：docs/python-excel/01-openpyxl.md（C15-L2-1 openpyxl）
# 运行方法：python3 Python-openpyxl-report.py
# 运行环境：Python 3.12，openpyxl 3.1.2（本机实测通过）
# 功能：用 openpyxl 生成一份带样式和公式的月度销售报表，
#       输出 Python-openpyxl-report.xlsx（明细表 + 按产品汇总表）。
# 数据：与阶段 6 教程共用同一份 12 行订单（总销售额 75600，数量合计 62）。
# 注意：金额列与合计行写入的是 Excel 公式（例如 =E2*F2），
#       openpyxl 本身不计算公式，公式结果需用 Excel 打开后计算；
#       本脚本在保存后重新打开文件核对公式文本，并用 Python 独立算出预期值。

import os

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# 订单数据：订单号、日期、客户编号、产品编号、数量、单价
ORDERS = [
    ("O001", "2025-01-05", "C001", "P001", 2, 5000),
    ("O002", "2025-01-12", "C002", "P002", 3, 3000),
    ("O003", "2025-02-08", "C001", "P003", 10, 300),
    ("O004", "2025-03-15", "C003", "P004", 5, 200),
    ("O005", "2025-03-20", "C002", "P005", 2, 1500),
    ("O006", "2025-06-10", "C004", "P001", 1, 5000),
    ("O007", "2026-01-06", "C001", "P001", 3, 5000),
    ("O008", "2026-01-15", "C002", "P002", 2, 3000),
    ("O009", "2026-02-10", "C003", "P003", 20, 300),
    ("O010", "2026-02-18", "C001", "P004", 8, 200),
    ("O011", "2026-03-05", "C004", "P005", 4, 1500),
    ("O012", "2026-03-22", "C002", "P001", 2, 5000),
]

PRODUCTS = {"P001": "电脑", "P002": "手机", "P003": "耳机", "P004": "键盘", "P005": "显示器"}

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
TITLE_FONT = Font(name="微软雅黑", size=14, bold=True, color="1F4E78")
TOTAL_FONT = Font(name="微软雅黑", size=11, bold=True)
THIN_BORDER = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
MONEY_FMT = '#,##0'
CENTER = Alignment(horizontal="center", vertical="center")


def style_header(ws, row, ncols):
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = THIN_BORDER


def build_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "明细"
    ws["A1"] = "销售明细（2025-01 至 2026-03）"
    ws["A1"].font = TITLE_FONT
    headers = ["订单号", "日期", "客户编号", "产品", "数量", "单价", "金额"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=2, column=col, value=h)
    style_header(ws, 2, len(headers))
    for i, (oid, date, cust, pid, qty, price) in enumerate(ORDERS, start=3):
        ws.cell(row=i, column=1, value=oid)
        ws.cell(row=i, column=2, value=date)
        ws.cell(row=i, column=3, value=cust)
        ws.cell(row=i, column=4, value=PRODUCTS[pid])
        ws.cell(row=i, column=5, value=qty)
        ws.cell(row=i, column=6, value=price).number_format = MONEY_FMT
        # 金额列写公式，不写死数字：改数量或单价后 Excel 重算即更新
        ws.cell(row=i, column=7, value=f"=E{i}*F{i}").number_format = MONEY_FMT
    last = 2 + len(ORDERS)
    ws.cell(row=last, column=1, value="合计").font = TOTAL_FONT
    ws.cell(row=last, column=5, value=f"=SUM(E3:E{last - 1})").font = TOTAL_FONT
    ws.cell(row=last, column=7, value=f"=SUM(G3:G{last - 1})").font = TOTAL_FONT
    ws.cell(row=last, column=7).number_format = MONEY_FMT
    for row in ws.iter_rows(min_row=3, max_row=last, max_col=7):
        for cell in row:
            cell.border = THIN_BORDER
    widths = {"A": 10, "B": 13, "C": 11, "D": 10, "E": 8, "F": 10, "G": 12}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:G{last - 1}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # 按产品汇总表：数值由 Python 算出，合计行用公式
    totals = {}
    for _oid, _date, _cust, pid, qty, price in ORDERS:
        name = PRODUCTS[pid]
        t = totals.setdefault(name, [0, 0, 0])
        t[0] += 1
        t[1] += qty
        t[2] += qty * price
    ws2 = wb.create_sheet("按产品汇总")
    ws2["A1"] = "按产品汇总"
    ws2["A1"].font = TITLE_FONT
    for col, h in enumerate(["产品", "订单数", "数量", "销售额"], start=1):
        ws2.cell(row=2, column=col, value=h)
    style_header(ws2, 2, 4)
    for i, name in enumerate(["电脑", "手机", "耳机", "键盘", "显示器"], start=3):
        n, qty, amount = totals[name]
        ws2.cell(row=i, column=1, value=name)
        ws2.cell(row=i, column=2, value=n)
        ws2.cell(row=i, column=3, value=qty)
        ws2.cell(row=i, column=4, value=amount).number_format = MONEY_FMT
    total_row = 8
    ws2.cell(row=total_row, column=1, value="合计").font = TOTAL_FONT
    for col, letter in ((2, "B"), (3, "C"), (4, "D")):
        ws2.cell(row=total_row, column=col, value=f"=SUM({letter}3:{letter}7)").font = TOTAL_FONT
    ws2.cell(row=total_row, column=4).number_format = MONEY_FMT
    for row in ws2.iter_rows(min_row=3, max_row=total_row, max_col=4):
        for cell in row:
            cell.border = THIN_BORDER
    for col, w in {"A": 10, "B": 10, "C": 10, "D": 12}.items():
        ws2.column_dimensions[col].width = w
    return wb


def main():
    out_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(out_dir, "Python-openpyxl-report.xlsx")
    wb = build_workbook()
    wb.save(out_path)

    # Python 独立验算（不依赖 Excel 公式计算结果）
    total_qty = sum(o[4] for o in ORDERS)
    total_amount = sum(o[4] * o[5] for o in ORDERS)
    print("订单数：", len(ORDERS))
    print("数量合计：", total_qty)
    print("销售额合计：", total_amount)

    # 重新打开文件核对：公式文本与汇总值
    wb2 = load_workbook(out_path)
    ws = wb2["明细"]
    print("明细 G3 公式：", ws["G3"].value)
    print("明细 G14 公式：", ws["G14"].value)
    print("明细 E14 公式：", ws["E14"].value)
    assert ws["G3"].value == "=E3*F3"
    assert ws["G14"].value == "=SUM(G3:G13)"
    assert ws["E14"].value == "=SUM(E3:E13)"
    assert ws["A2"].value == "订单号" and ws.max_row == 14
    ws2 = wb2["按产品汇总"]
    assert ws2["D8"].value == "=SUM(D3:D7)"
    assert ws2["D3"].value == 40000  # 电脑
    assert ws2["D7"].value == 9000   # 显示器
    assert total_qty == 62 and total_amount == 75600
    print("核对通过：", out_path)


if __name__ == "__main__":
    main()
