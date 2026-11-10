# Python-pandas-analysis.py
# 对应教程：docs/python-excel/02-pandas.md（C15-L2-2 pandas）
# 运行方法：python3 Python-pandas-analysis.py
# 运行环境：Python 3.12，pandas 2.1.4，openpyxl 3.1.2（本机实测通过）
# 功能：读入订单数据，做数据清洗、分组汇总与利润分析，
#       输出 Python-pandas-analysis.xlsx（原始数据 / 清洗后明细 /
#       按产品汇总 / 按地区汇总 / 按月汇总，共 5 张表）。
# 数据：与阶段 6 教程共用同一份 12 行订单（总销售额 75600，数量合计 62，
#       总利润 22770，利润率 30.12%）；另加 3 行脏数据演示清洗。
# 校验：分组汇总数与手工计算一致（电脑 40000、华东 29600、2026-01 月 21000）。

import os

import pandas as pd

# 干净的 12 行订单：订单号、日期、客户编号、产品编号、数量、单价
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

PRODUCTS = pd.DataFrame([
    {"产品编号": "P001", "产品": "电脑", "单价": 5000, "成本": 3800},
    {"产品编号": "P002", "产品": "手机", "单价": 3000, "成本": 2100},
    {"产品编号": "P003", "产品": "耳机", "单价": 300, "成本": 150},
    {"产品编号": "P004", "产品": "键盘", "单价": 200, "成本": 110},
    {"产品编号": "P005", "产品": "显示器", "单价": 1500, "成本": 1000},
])

CUSTOMERS = pd.DataFrame([
    {"客户编号": "C001", "客户": "华东公司", "地区": "华东"},
    {"客户编号": "C002", "客户": "华南公司", "地区": "华南"},
    {"客户编号": "C003", "客户": "华北公司", "地区": "华北"},
    {"客户编号": "C004", "客户": "华西公司", "地区": "华西"},
])

# 3 行脏数据：重复行、缺失单价、无法识别的单价文本（清洗后应恰好回到 12 行）
DIRTY_EXTRA = [
    ("O003", "2025-02-08", "C001", "P003", 10, 300),   # 与 O003 完全重复
    ("O013", "2026-03-25", "C002", "P001", 1, None),    # 单价缺失
    ("O014", "2026-03-28", "C003", "P005", 1, "待确认"),  # 单价是无法转数值的文本
]


def main():
    out_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(out_dir, "Python-pandas-analysis.xlsx")

    raw = pd.DataFrame(
        ORDERS + DIRTY_EXTRA,
        columns=["订单号", "日期", "客户编号", "产品编号", "数量", "单价"],
    )
    print("原始行数：", len(raw))

    # 清洗：文本单价转数值 -> 去重 -> 删缺失 -> 日期解析 -> 算金额
    df = raw.copy()
    df["单价"] = pd.to_numeric(df["单价"], errors="coerce")
    df = df.drop_duplicates()
    df = df.dropna(subset=["单价"])
    df["日期"] = pd.to_datetime(df["日期"])
    df["金额"] = df["数量"] * df["单价"]
    df = df.merge(PRODUCTS[["产品编号", "产品", "成本"]], on="产品编号")
    df = df.merge(CUSTOMERS[["客户编号", "地区"]], on="客户编号")
    df["利润"] = df["数量"] * (df["单价"] - df["成本"])
    df["月份"] = df["日期"].dt.strftime("%Y-%m")
    print("清洗后行数：", len(df))
    print("销售额合计：", int(df["金额"].sum()))
    print("数量合计：", int(df["数量"].sum()))
    print("利润合计：", int(df["利润"].sum()))
    print("利润率：", round(df["利润"].sum() / df["金额"].sum() * 100, 2), "%")

    by_product = df.groupby("产品").agg(订单数=("订单号", "count"), 数量=("数量", "sum"),
                                        销售额=("金额", "sum"), 利润=("利润", "sum")).reset_index()
    by_region = df.groupby("地区").agg(订单数=("订单号", "count"), 数量=("数量", "sum"),
                                       销售额=("金额", "sum")).reset_index()
    by_month = df.groupby("月份").agg(订单数=("订单号", "count"), 销售额=("金额", "sum")).reset_index()
    print(by_product.to_string(index=False))
    print(by_region.to_string(index=False))
    print(by_month.to_string(index=False))

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        raw.to_excel(writer, sheet_name="原始数据", index=False)
        df.to_excel(writer, sheet_name="清洗后明细", index=False)
        by_product.to_excel(writer, sheet_name="按产品汇总", index=False)
        by_region.to_excel(writer, sheet_name="按地区汇总", index=False)
        by_month.to_excel(writer, sheet_name="按月汇总", index=False)

    # 重新读回文件核对：分组汇总与手工计算一致
    check = pd.read_excel(out_path, sheet_name="按产品汇总")
    assert len(df) == 12 and int(df["金额"].sum()) == 75600
    assert int(df["数量"].sum()) == 62 and int(df["利润"].sum()) == 22770
    assert int(check.loc[check["产品"] == "电脑", "销售额"].iloc[0]) == 40000
    assert int(check.loc[check["产品"] == "键盘", "销售额"].iloc[0]) == 2600
    check_r = pd.read_excel(out_path, sheet_name="按地区汇总")
    assert int(check_r.loc[check_r["地区"] == "华东", "销售额"].iloc[0]) == 29600
    check_m = pd.read_excel(out_path, sheet_name="按月汇总")
    assert int(check_m.loc[check_m["月份"] == "2026-01", "销售额"].iloc[0]) == 21000
    print("核对通过：", out_path)


if __name__ == "__main__":
    main()
