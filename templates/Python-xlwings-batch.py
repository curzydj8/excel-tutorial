# Python-xlwings-batch.py
# 对应教程：docs/python-excel/03-xlwings.md（C15-L2-3 xlwings）
# 运行方法：python3 Python-xlwings-batch.py
# 运行环境：Python 3.12，xlwings 0.37.8（本机实测 import 通过）
# 功能分两段：
#   (1) 可执行部分：用 openpyxl 生成待合并的分表样本
#       Python-xlwings-sample.xlsx（华东 / 华南 / 华西 3 张分表），
#       并做 import 与语法检查；
#   (2) Excel 驱动部分（函数 merge_branch_reports）：调用 xw.App / xw.Book，
#       把 3 张分表合并为一张汇总表。未在 Excel 内执行验证：
#       本机为 Linux 且无 Excel，xlwings 交互模式仅支持 Windows 与 macOS，
#       实测 xw.apps.keys() 报错、xw.App() 报错（见教程 03 篇第 5 节）。
#       请在装有 Excel 的 Windows 或 macOS 上运行本文件第 (2) 段。

import os

import xlwings as xw
from openpyxl import Workbook

print("xlwings 版本：", xw.__version__)

# 3 个分公司的分表数据：分公司、订单号、产品、数量、单价
BRANCH_DATA = {
    "华东": [("O001", "电脑", 2, 5000), ("O007", "电脑", 3, 5000)],
    "华南": [("O002", "手机", 3, 3000), ("O008", "手机", 2, 3000)],
    "华西": [("O006", "电脑", 1, 5000), ("O011", "显示器", 4, 1500)],
}


def build_sample():
    """纯 Python 部分：生成 3 张分表样本（不依赖 Excel，可在本机运行）。"""
    out_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(out_dir, "Python-xlwings-sample.xlsx")
    wb = Workbook()
    first = True
    for branch, rows in BRANCH_DATA.items():
        ws = wb.active if first else wb.create_sheet(branch)
        ws.title = branch
        first = False
        ws.append(["订单号", "产品", "数量", "单价"])
        for row in rows:
            ws.append(list(row))
    wb.save(out_path)
    print("分表样本已生成：", out_path)
    return out_path


def merge_branch_reports(sample_path, output_path):
    """Excel 驱动部分：未在 Excel 内执行验证（本机无 Excel，无法运行）。"""
    # 未在 Excel 内执行验证：以下 xw.App / xw.Book 调用需在 Windows 或 macOS 上执行。
    app = xw.App(visible=False)
    try:
        book = app.books.open(sample_path)
        summary = book.sheets.add("汇总")
        summary["A1"].value = ["分公司", "订单号", "产品", "数量", "单价", "金额"]
        row = 2
        for sheet in book.sheets:
            if sheet.name == "汇总":
                continue
            data = sheet["A2"].expand("table").value
            for order_id, product, qty, price in data:
                summary[row, 1].value = [sheet.name, order_id, product, qty, price, qty * price]
                row += 1
        summary[f"A{row}"].value = "合计"
        summary[f"E{row}"].value = f"=SUM(E2:E{row - 1})"
        summary[f"F{row}"].value = f"=SUM(F2:F{row - 1})"
        book.save(output_path)
        print("汇总完成：", output_path)
    finally:
        app.quit()


def main():
    sample_path = build_sample()
    try:
        # 本机无 Excel：此调用会报错，仅用于如实报告环境限制，不视为已验证。
        merge_branch_reports(sample_path, "Python-xlwings-merged.xlsx")
    except Exception as exc:  # noqa: BLE001 - 此处故意捕获全部异常并如实报告
        print("Excel 驱动部分未执行：", type(exc).__name__, str(exc)[:160])
        print("原因：本机为 Linux 且无 Excel；请在 Windows 或 macOS 上重跑本脚本。")


if __name__ == "__main__":
    main()
