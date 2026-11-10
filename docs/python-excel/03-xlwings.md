# xlwings：调用正在运行的 Excel，批量处理工作簿

- **所属栏目**：C15 Python 与 Excel ／ C15-L2-3 xlwings 与自动化
- **适用版本**：Python 3.10 及以上，xlwings 0.37.8；Excel 驱动部分需 Windows 或 macOS 上安装的 Excel（Microsoft 365 或买断版均可，xlwings 不依赖特定 Excel 版本）
- **难度**：进阶
- **预计阅读时长**：40 分钟
- **更新日期**：2026-11-10

## 1. 学习目标

学完本篇，你将能够：
- 说出 xlwings 与 openpyxl 的根本区别：前者驱动正在运行的 Excel，后者直接读写文件
- 安装 xlwings 并检查版本，说清本机为什么跑不了 Excel 驱动代码
- 用 xlwings 打开工作簿、读写单元格区域、调用 Excel 的公式重算与保存
- 写一个批量处理脚本：把多个分表合并成一张汇总表
- 判断一个任务该用 xlwings 还是 openpyxl（或 pandas），并说出理由

**验证口径（必读）：** 本篇分两类代码。第一类是纯 Python 代码（import 检查、环境检查、用 openpyxl 准备分表样本、合并结果的 Python 侧验算），已在本机真实运行通过，运行结果为真实运行输出。第二类是 Excel 驱动代码（`xw.App`、`xw.Book` 等调用），**未在 Excel 内执行验证**：本机为 Linux 且无 Excel，而 xlwings 交互模式仅支持 Windows 与 macOS；本机实测 `xw.apps.keys()` 报错、`xw.App()` 报错（见示例 1 的真实输出）。每段 Excel 驱动代码处均已标注“未在 Excel 内执行验证”，请在装有 Excel 的 Windows 或 macOS 上运行。配套可运行脚本见第 11 节。

## 2. 基础概念

### 2.1 xlwings 是什么

xlwings 是一个让 Python 驱动 Excel 的库：Python 脚本通过 xlwings 向正在运行的 Excel 发命令——打开文件、选中区域、写入数值、调用宏、触发重算，再把结果读回来。Excel 是执行者，Python 是指挥者。

### 2.2 xlwings 与 openpyxl 的根本区别

| 对比项 | openpyxl（01 篇） | xlwings（本篇） |
|---|---|---|
| 需不需要装 Excel | 不需要，在 Linux 服务器上也能跑 | 需要，目标机器上必须装有 Excel |
| 支持的操作系统 | Windows、macOS、Linux 都行 | 交互模式仅 Windows 与 macOS（Linux 上只能用 xlwings Server 或 Reader） |
| 工作方式 | 直接读写 xlsx 文件 | 驱动正在运行的 Excel 程序 |
| 公式 | 只写公式文本，不计算 | 可触发 Excel 重算，读到计算结果 |
| 宏与加载项 | 不能运行 | 可以调用 VBA 宏 |
| 速度 | 快，纯文件操作 | 慢，每步都要经过 Excel |
| 适用场景 | 服务器上批量生成报表 | 办公室里批量操作已打开的工作簿、调用宏、做 Excel 界面做不到的自动化 |

一句话：文件不在 Excel 里打开、也不需要 Excel 算公式时用 openpyxl；反之，任务必须经过 Excel 本身时用 xlwings。

### 2.3 四个核心对象

- 应用（App）：`xw.App(visible=False)` 启动一个 Excel 进程；`visible=False` 表示后台运行不显示窗口。
- 工作簿（Book）：`app.books.open(路径)` 打开一个文件；`app.books.add()` 新建一个。
- 工作表（Sheet）：`book.sheets[表名]` 或 `book.sheets[0]` 取表；`book.sheets.add(表名)` 新建表。
- 区域（Range）：`sheet["A1"]` 取一个格子，`sheet["A1:C5"]` 取一个区域，`.value` 读写值；`sheet["A2"].expand("table")` 从 A2 起扩展到整张连续数据表。

### 2.4 安装

```bash
pip install xlwings
```

注意：在 Debian／Ubuntu 系的 Python 上直接 `pip install xlwings` 可能报外部环境管理错误（PEP 668），本机实测就是如此；解决办法是用虚拟环境，或加 `--break-system-packages` 参数。装好后检查版本（本机实测版本 0.37.8）：

```python
import xlwings as xw
print(xw.__version__)
```

运行结果：

```
0.37.8
```

## 3. 操作步骤

1. 检查环境：确认 xlwings 能 import，确认本机有没有可驱动的 Excel（示例 1；本机没有，往下走准备步骤）。
2. 准备分表样本：用 openpyxl 生成 3 张分公司分表（华东、华南、华西），作为批量合并的输入（示例 2；纯 Python，可在本机运行）。
3. 打开工作簿：`xw.App(visible=False)` 启动 Excel，`app.books.open(路径)` 打开分表文件（示例 3；未在 Excel 内执行验证）。
4. 批量合并：循环 3 张分表逐行读出，写入“汇总”表，金额列写公式，末尾加合计行（示例 4；未在 Excel 内执行验证）。
5. Python 侧验算：不用 Excel，用 openpyxl 读分表样本，在 Python 里算出合并后的预期结果（6 行、数量 15、金额 51000），供在 Windows 或 macOS 上运行后对照（示例 5；可运行）。

## 4. 图文示例

示例：3 张分表合并成 1 张汇总表（文字版）。

分表（每个分公司一张表，结构相同）：

| 订单号 | 产品 | 数量 | 单价 |
|---|---|---|---|
| O001 | 电脑 | 2 | 5000 |
| O007 | 电脑 | 3 | 5000 |

华东 2 行、华南 2 行（O002 手机 3 件、O008 手机 2 件）、华西 2 行（O006 电脑 1 件、O011 显示器 4 件），共 6 行。

合并后的汇总表：

| 分公司 | 订单号 | 产品 | 数量 | 单价 | 金额 |
|---|---|---|---|---|---|
| 华东 | O001 | 电脑 | 2 | 5000 | 10000 |
| 华东 | O007 | 电脑 | 3 | 5000 | 15000 |
| 华南 | O002 | 手机 | 3 | 3000 | 9000 |
| 华南 | O008 | 手机 | 2 | 3000 | 6000 |
| 华西 | O006 | 电脑 | 1 | 5000 | 5000 |
| 华西 | O011 | 显示器 | 4 | 1500 | 6000 |
| 合计 | | | 15 | | 51000 |

数量合计 2+3+3+2+1+4 = 15，金额合计 10000+15000+9000+6000+5000+6000 = 51000（待补图）。

## 5. 公式和代码

示例 1：import 检查与环境检查（可运行）。

```python
import xlwings as xw

print("xlwings 版本：", xw.__version__)
try:
    print("正在运行的 Excel：", xw.apps.keys())
except Exception as exc:
    print("检查 Excel 失败：", type(exc).__name__, str(exc))
try:
    app = xw.App(visible=False)
    print("Excel 已启动")
    app.quit()
except Exception as exc:
    print("启动 Excel 失败：", type(exc).__name__, str(exc)[:120])
```

运行结果（本机真实运行输出，Linux，无 Excel）：

```
xlwings 版本： 0.37.8
检查 Excel 失败： XlwingsError The interactive mode of xlwings is only supported on Windows and macOS. On Linux, you can use xlwings Server or xlwings Reader.
启动 Excel 失败： AttributeError 'NoneType' object has no attribute 'apps'
```

结论：xlwings 本身装好了（0.37.8），但本机没有 Excel 可驱动。报错信息也说得很清楚：交互模式只支持 Windows 与 macOS。在 Windows 或 macOS 上跑同一段代码，`xw.apps.keys()` 会返回正在运行的 Excel 进程列表，`xw.App(visible=False)` 会以后台方式启动 Excel。

示例 2：用 openpyxl 准备分表样本（可运行，纯 Python 部分）。

```python
from openpyxl import Workbook

BRANCH_DATA = {
    "华东": [("O001", "电脑", 2, 5000), ("O007", "电脑", 3, 5000)],
    "华南": [("O002", "手机", 3, 3000), ("O008", "手机", 2, 3000)],
    "华西": [("O006", "电脑", 1, 5000), ("O011", "显示器", 4, 1500)],
}
wb = Workbook()
first = True
for branch, rows in BRANCH_DATA.items():
    ws = wb.active if first else wb.create_sheet(branch)
    ws.title = branch
    first = False
    ws.append(["订单号", "产品", "数量", "单价"])
    for row in rows:
        ws.append(list(row))
wb.save("分表样本.xlsx")
print("工作表：", wb.sheetnames)
print("每表行数：", [wb[name].max_row for name in wb.sheetnames])
```

运行结果（本机真实运行输出）：

```
工作表： ['华东', '华南', '华西']
每表行数： [3, 3, 3]
```

每张表 3 行 = 1 行标题 + 2 行数据。这个文件就是示例 3 和示例 4 的输入；准备数据的活不需要 Excel，用 openpyxl 做更合适——这正是分工的一部分。

示例 3：打开工作簿、读写单元格（未在 Excel 内执行验证）。

```python
import xlwings as xw

# 未在 Excel 内执行验证：以下代码需在装有 Excel 的 Windows 或 macOS 上运行。
app = xw.App(visible=False)
book = app.books.open("分表样本.xlsx")
sheet = book.sheets["华东"]
print(sheet["A1"].value)          # 读一个格子：订单号
print(sheet["A1"].expand("table").value)  # 从 A1 扩展读整张表（含标题行）
print(sheet["A1"].expand("table").value[1:])  # 去掉标题行，只剩数据行
sheet["A5"].value = "O099"        # 写一个格子
book.save()
book.close()
app.quit()
```

预期结果（未在 Excel 内执行验证）：`sheet["A1"].value` 是 `订单号`；`expand("table")` 从 A1 起扩展读到的整表是 `[["订单号", "产品", "数量", "单价"], ["O001", "电脑", 2, 5000], ["O007", "电脑", 3, 5000]]`，其中第一行是标题行；切片 `[1:]` 去掉标题行后只剩两行数据。注意：`expand` 扩展的是包含起始格的整片连续区域，从 A2 起扩展同样会把上方的标题行包含进来，所以读数据行时要记得切掉第一行。要点：`app.quit()` 一定要写，否则 Excel 进程留在后台，下次再跑会打开第二个 Excel；读写完先 `book.save()` 再 `book.close()`。

示例 4：批量合并 3 张分表（未在 Excel 内执行验证）。

```python
import xlwings as xw

# 未在 Excel 内执行验证：以下代码需在装有 Excel 的 Windows 或 macOS 上运行。
app = xw.App(visible=False)
try:
    book = app.books.open("分表样本.xlsx")
    summary = book.sheets.add("汇总")
    summary["A1"].value = ["分公司", "订单号", "产品", "数量", "单价", "金额"]
    row = 2
    for sheet in book.sheets:
        if sheet.name == "汇总":
            continue
        data = sheet["A1"].expand("table").value[1:]  # 去掉标题行，只剩数据行
        for order_id, product, qty, price in data:
            summary[row, 1].value = [sheet.name, order_id, product, qty, price, qty * price]
            row += 1
    summary[f"A{row}"].value = "合计"
    summary[f"E{row}"].value = f"=SUM(E2:E{row - 1})"
    summary[f"F{row}"].value = f"=SUM(F2:F{row - 1})"
    book.save("合并汇总.xlsx")
    print("汇总完成，共写入数据行：", row - 2)
finally:
    app.quit()
```

预期结果（未在 Excel 内执行验证）：汇总表 8 行（1 行标题 + 6 行数据 + 1 行合计），数量合计 15，金额合计 51000（见示例 5 的 Python 侧验算）。要点：循环里用 `sheet.name == "汇总"` 跳过新加的汇总表，否则会把汇总表自己也当分表读一遍；`try...finally` 保证即使中途报错也会 `app.quit()`，不留后台 Excel 进程；`summary[row, 1].value = [...]` 一次写一整行，比逐格写快得多。

示例 5：合并结果的 Python 侧验算（可运行，不用 Excel）。

```python
from openpyxl import load_workbook

wb = load_workbook("分表样本.xlsx")
total_qty = 0
total_amount = 0
rows = 0
for name in ["华东", "华南", "华西"]:
    for _oid, _product, qty, price in wb[name].iter_rows(min_row=2, values_only=True):
        rows += 1
        total_qty += qty
        total_amount += qty * price
print("合并后数据行数：", rows)
print("数量合计：", total_qty)
print("金额合计：", total_amount)
assert rows == 6 and total_qty == 15 and total_amount == 51000
print("验算通过：与第 4 节汇总表一致")
```

运行结果（本机真实运行输出）：

```
合并后数据行数： 6
数量合计： 15
金额合计： 51000
验算通过：与第 4 节汇总表一致
```

这个验算的意义：在 Windows 或 macOS 上跑完示例 4 后，打开“合并汇总.xlsx”核对合计行——数量应为 15、金额应为 51000；对不上就说明合并循环漏读或重复读了某张表。

## 6. 实际工作场景

场景 1：每月合并 30 个分公司的报表。

- 背景：30 个分公司每月各交一张 xlsx，表头相同、行数不同，要合成一张总表再汇总；以前的做法是逐个打开复制粘贴，半天过去还容易粘错行。
- 做法：按示例 4 的脚本循环：`app.books.open` 打开一个分表，`expand("table")` 读整表，逐行追加到汇总表，处理完关闭该分表；30 个文件跑完后在汇总表末尾加合计行公式，一次保存。
- 结果：以本篇 3 张分表为例，合并后 6 行、数量 15、金额 51000；30 张表同理，只是循环次数变多。注意：每处理完一张分表就 `close`，不要同时打开 30 个工作簿。

场景 2：调用 VBA 宏完成 Excel 界面做不了的收尾。

- 背景：报表生成后还要执行一段现成的 VBA 宏（例如按客户拆分工作表并分别打印设置），宏已经在文件里，手工点一次要等很久。
- 做法：xlwings 打开文件后调 `book.macro("宏名")()` 执行宏，再 `book.save()` 保存；Python 负责循环与文件管理，宏负责 Excel 内部操作。
- 结果：Python 与 VBA 各做擅长的事；这个场景 openpyxl 做不了，因为它不能运行宏——这也是 xlwings 不可替代的一处。

## 7. 常见错误

| 错误现象 | 原因 | 解决方案 |
|---|---|---|
| `XlwingsError: The interactive mode of xlwings is only supported on Windows and macOS` | 在 Linux 上调用 `xw.App()` 或 `xw.apps` | 交互模式只支持 Windows 与 macOS；Linux 上改用 openpyxl／pandas，或用 xlwings Server／Reader 方案 |
| `AttributeError: 'NoneType' object has no attribute 'apps'` | 同上：Linux 上没有 Excel 进程可连 | 不是代码写错，是环境不支持；换一台装有 Excel 的 Windows 或 macOS 再跑 |
| 脚本跑完后任务管理器里一堆 EXCEL.EXE | 只 `open` 不 `quit`，Excel 进程留在后台 | 每次都用 `try...finally: app.quit()`；已残留的进程在任务管理器里手动结束 |
| `book.save()` 报文件被占用 | 该文件正在 Excel 里手动打开着 | 先关掉手动打开的文件，或另存为新文件名；不要让脚本与手工同时写同一个文件 |
| `expand("table")` 读到的表缺行或多行 | 数据中间有空行或空列，扩展在空行处停下 | 保证分表数据连续无空行；或不用 `expand`，改用 `sheet.used_range.value` 整片读再处理 |
| 写入的值在 Excel 里显示为文本（如 5000 左对齐） | 写入的是字符串 `"5000"` 而不是数值 `5000` | 写入前用 `int()`／`float()` 转数值；读分表时注意 `expand` 读到的文本型数字 |
| 循环写几千行非常慢 | 逐格写 `sheet["A1"].value`，每次写都是一次 Excel 调用 | 一次写一整行（`summary[row, 1].value = [...]`）或整片写（`range.value = 二维列表`） |
| 宏调用报 `book.macro` 找不到宏 | 宏名写错，或宏在另一个工作簿／加载宏里 | 核对宏名与所在模块；宏在 PERSONAL.XLSB 时先确认该文件已加载 |
| `visible=False` 后 Excel 窗口还是弹出来了 | 已有一个可见的 Excel 进程在运行，xlwings 连上了它 | 先关掉已打开的 Excel，或用 `xw.App(visible=False, add_book=False)` 新起一个后台进程 |
| 中文表名或中文路径报编码错误 | 旧版 xlwings 对非 ASCII 路径处理问题，或文件路径含特殊字符 | 升级 xlwings 到最新版；路径用英文文件夹过渡，或用 `os.path.abspath` 取绝对路径再打开 |

## 8. 相关教程

- 同栏目上一篇：02-pandas.md（pandas：DataFrame 与 Excel 互转、数据清洗、分组汇总）
- 同栏目下一篇：04-python-in-excel.md（Python in Excel：=PY() 入门、适用与限制）
- 同栏目相关：01-openpyxl.md（openpyxl：用 Python 读写工作簿；与本篇的分工对照见第 2.2 节）
- 跨栏目相关：docs/vba/04-examples.md（VBA 实战实例：批量合并工作簿等 8 个实例，与本篇示例 4 对照）
- 跨栏目相关：docs/vba/01-basics.md（VBA 基础语法：示例 4 调用的宏所用的语言）
- 跨栏目相关：docs/powerquery/01-power-query.md（Power Query：不用写代码合并多文件，适合与本篇对照选型）

## 9. 练习题

1. 在本机运行 `import xlwings as xw; print(xw.__version__)`，报出版本号；再运行示例 1 的环境检查，写出两条报错信息分别说明了什么。
2. 用 openpyxl 生成 3 张分表样本（华东、华南、华西各 2 行），保存后重新打开，打印每张表的工作表名与行数。
3. 说出 `xw.App(visible=False)` 里的 `visible=False` 是什么意思；如果改成 `visible=True`，运行中会看到什么不同？
4. 示例 4 的循环里为什么要用 `if sheet.name == "汇总": continue` 跳过汇总表？如果不跳过，会发生什么？
5. 写出调用 VBA 宏的 xlwings 代码（假设宏名叫“拆分并打印设置”），并说出 `book.save()` 应该写在调用宏之前还是之后，为什么？
6. 判断题：以下三个任务各该用 xlwings、openpyxl 还是 pandas？（1）在 Linux 服务器上每晚定时生成 100 份报表；（2）在 Windows 办公室里把 30 张已打开的分表合并，并调用现成的 VBA 宏收尾；（3）一份 10 万行订单要清洗、分组汇总并输出报告。每个判断写出一句话理由。

## 10. 参考答案

<details>
<summary>点击查看答案</summary>

1. 本机实测版本 0.37.8。两条报错：`XlwingsError: The interactive mode ... only supported on Windows and macOS` 说明交互模式只支持 Windows 与 macOS，Linux 上没有可驱动的 Excel；`AttributeError: 'NoneType' object has no attribute 'apps'` 是同一原因的连带报错，`xw.App()` 在 Linux 上拿不到 Excel 应用对象。两条都不是代码写错，而是环境不支持。
2. 做法见示例 2。重新打开后工作表名是 `['华东', '华南', '华西']`，每表行数 3（1 行标题 + 2 行数据）。如果某表行数是 2，说明漏写了标题行。
3. `visible=False` 表示以后台方式启动 Excel，不显示窗口；改成 `visible=True` 会看到 Excel 窗口弹出来并实时执行每一步，适合调试，跑批量任务时用 False 更快且不干扰操作。
4. 因为汇总表是循环中新加的表，不跳过的话，循环走到汇总表时会把汇总表里刚写入的行又当作分表数据读一遍、再追加进去，造成重复行与合计越算越大。
5. 代码：`book.macro("拆分并打印设置")()`。`book.save()` 应写在调用宏之后：宏会改动工作簿内容（拆分表、改打印设置），保存要在所有改动完成后做；先保存再调宏，宏做的改动就没存进去。
6. （1）openpyxl：Linux 服务器上没有 Excel，xlwings 跑不了；定时生成报表是纯文件操作，openpyxl 最合适。（2）xlwings：任务必须经过正在运行的 Excel，还要调用现成的 VBA 宏，只有 xlwings 能做到。（3）pandas：10 万行清洗与分组汇总是 pandas 的主场，`groupby` 一行代码完成；openpyxl 逐格写 10 万行太慢，xlwings 更慢。

</details>

## 11. 可下载示例文件

- Python-xlwings-batch.py（templates/Python-xlwings-batch.py，xlwings 批量合并脚本：纯 Python 部分可直接运行，Excel 驱动部分需在 Windows 或 macOS 上运行，运行方法 `python3 Python-xlwings-batch.py`）
- Python-xlwings-sample.xlsx（templates/Python-xlwings-sample.xlsx，脚本运行产物：华东 / 华南 / 华西 3 张分表，每表 2 行数据，供合并脚本读取）
