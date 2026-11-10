# Python-in-excel-demo.py
# 对应教程：docs/python-excel/04-python-in-excel.md（C15-L2-4 Python in Excel）
# 运行方法：python3 Python-in-excel-demo.py
# 运行环境：Python 3.12，pandas 2.1.4，openpyxl 3.1.2（本机实测通过）
# 功能：本文件是教程中 =PY() 示例的本地等价运算。
#       =PY() 代码在 Excel 云端运行，本机无法执行（未在 Excel 内执行验证）；
#       这里用同样的 pandas 逻辑在本地跑出同一批数字，
#       供读者在没有 Microsoft 365 的电脑上对照结果，
#       输出 Python-in-excel-demo.xlsx（明细 / 按产品汇总 / 描述统计）。
# 数据：与阶段 6 教程共用同一份 12 行订单（总销售额 75600，数量合计 62）。

import os

import pandas as pd

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


def main():
    out_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(out_dir, "Python-in-excel-demo.xlsx")

    # 与教程 =PY() 示例 2 相同的逻辑：读表 -> 算金额 -> 分组汇总
    df = pd.DataFrame(ORDERS, columns=["订单号", "日期", "客户编号", "产品编号", "数量", "单价"])
    df["产品"] = df["产品编号"].map(PRODUCTS)
    df["金额"] = df["数量"] * df["单价"]
    print("总销售额：", int(df["金额"].sum()))
    print("订单笔数：", len(df))
    print("客单价：", int(df["金额"].sum() // len(df)))

    # 与教程 =PY() 示例 3 相同的逻辑：describe() 描述统计
    desc = df[["数量", "单价", "金额"]].describe()
    print(desc.to_string())

    by_product = df.groupby("产品").agg(订单数=("订单号", "count"), 数量=("数量", "sum"),
                                        销售额=("金额", "sum")).reset_index()
    print(by_product.to_string(index=False))

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="明细", index=False)
        by_product.to_excel(writer, sheet_name="按产品汇总", index=False)
        desc.to_excel(writer, sheet_name="描述统计")

    check = pd.read_excel(out_path, sheet_name="按产品汇总")
    assert int(df["金额"].sum()) == 75600
    assert int(check.loc[check["产品"] == "电脑", "销售额"].iloc[0]) == 40000
    print("核对通过：", out_path)


if __name__ == "__main__":
    main()
