# M 语言入门：读懂 Power Query 背后的代码

- **所属栏目**：C09 Power Query ／ C09-L2-4 M 语言入门
- **适用版本**：Excel 2016／2019／2021／365（内置，位于“数据”选项卡）；Excel 2013／2010 需单独安装 Power Query 插件
- **难度**：高级
- **预计阅读时长**：40 分钟
- **更新日期**：2026-11-10

## 1. 学习目标

学完本篇，你将能够：
- 读懂 let/in 结构的 M 代码，用 each、Table. 与 Text. 常用函数写自定义列
- 在高级编辑器里修改查询，并把文件路径变成函数参数
- 说出 M 语言的三条基本规则（大小写、双引号、列表从 0 开始数）

## 2. 基础概念

### 2.1 你点的每一步都会变成一行 M 代码

你点的每一步，Power Query 都翻译成 M 语言，显示在顶部的公式栏。看懂它，就能做界面点不出来的事：左半联接、自定义函数、动态路径。先看、再改、最后自己写，是学 M 的顺序。

### 2.2 let/in：定义一串中间变量

`let...in` 是 M 最常见的结构：定义一串中间变量，最后 `in` 后面的是输出。每个步骤名（如“提升标题”）都是一个变量，下一步引用上一步。读代码时顺着变量名往下读，就是查询的执行顺序。

### 2.3 each：对每一行做判断

`each [金额] > 0` 是对每一行判断，`[金额]` 表示当前行的“金额”字段。`each` 是 `(_) => _` 的简写，相当于 M 里的 lambda。自定义列里几乎全是 `each` 开头。

### 2.4 列表、记录、表三种结构

M 的列表、记录、表三种结构是核心，掌握 `each` 和 `[字段名]` 就能写 80% 的自定义步骤：

| 结构 | 写法 | 取值 |
|---|---|---|
| 列表 (List) | `{"a", "b", "c"}` | `列表{1}` 得到 (b) |
| 记录 (Record) | `[姓名="王芳", 金额=9000]` | `记录[金额]` 得到 9000 |
| 表 (Table) | 多行多列 | `表{0}` 得到第 1 行记录，`表[金额]` 得到金额列的列表 |

### 2.5 三条基本规则

读代码时记住三条 M 语言的基本规则：M 区分大小写（Table.SelectRows 不能写成 table.selectrows）；文本用双引号，文本里的双引号要写两遍；列表从 0 开始数（{1} 是第 2 个元素）。

## 3. 操作步骤

1. 点查询进编辑器，点任意步骤，看顶部公式栏里的 M 代码。
2. 读懂代码：顺着 let/in 的变量名往下读，每步对应界面上的一个操作。
3. 改代码：点“主页→高级编辑器”，改一行（如把筛选条件从 `> 0` 改成 `> 100`），点“完成”。
4. 写自定义列：“添加列→自定义列”，用 `each` 写公式，如 `each if [金额] >= 10000 then "大" else "小"`。
5. 建函数：把文件路径变成参数（见 5.6 节），调用时传路径。
6. 出错时看公式栏：报错的步骤公式栏里会标红，先看引用的列名对不对。

## 4. 图文示例

### 4.1 读一段代码：从 CSV 到筛选

```m
let
    源 = Csv.Document(File.Contents("C:\data\销售.csv"), [Delimiter=",", Encoding=65001]),
    提升标题 = Table.PromoteHeaders(源),
    改类型 = Table.TransformColumnTypes(提升标题, {{"日期", type date}, {"金额", type number}}),
    筛选 = Table.SelectRows(改类型, each [金额] > 0)
in
    筛选
```

读法：`Csv.Document` 的第 2 个参数是选项记录：`[Delimiter=",", Encoding=65001]` 表示逗号分隔、UTF-8 编码（65001 是 UTF-8 的代码页）。`Table.TransformColumnTypes` 的第 2 个参数是列表，每个元素是（列名，类型）的二元组：`{"日期", type date}`。预期输出：读入销售 csv，首行提升为标题，“日期”列为日期型、“金额”列为数字型，只保留金额大于 0 的行。

### 4.2 自定义列：等级与城市

```m
// 新增条件列（界面：添加列 → 条件列 点不出的复杂逻辑自己写）
Table.AddColumn(上一步, "等级", each if [金额] >= 10000 then "大" else "小")

// 按分隔符取第 2 段
Table.AddColumn(上一步, "城市", each Text.Split([地址], "-"){1})
```

`{1}` 是取列表第 2 个元素（从 0 开始数）。`Text.Split("广东-深圳-南山区", "-")` 返回列表 `{"广东", "深圳", "南山区"}`，取 `{1}` 得到（深圳）。预期输出（以地址“广东-深圳-南山区”为例）：等级列按金额分（大/小），城市列为（深圳）。

![M 语言入门：读懂 Power Query 背后](images/u12-powerquery-04-m-language-01.png)

## 5. 公式和代码

### 5.0 验证状态说明（必读）

本节全部 M 代码的验证状态：**未执行验证（无 Power Query 运行时）**。本机没有可运行 Power Query 的 Excel 环境，因此没有任何一段代码声称已执行成功。

已做的核对只有两项：

- **语法经官方文档核对**：每个函数名与参数签名均对照 Microsoft Learn 的 Power Query M 函数参考核对过（Csv.Document、File.Contents、Folder.Files、Excel.Workbook、Sql.Database、Web.Contents、Json.Document、Table.PromoteHeaders、Table.TransformColumnTypes、Table.SelectRows、Table.AddColumn、Table.SplitColumn、Splitter.SplitTextByDelimiter、Table.FillDown、Table.NestedJoin、Table.ExpandTableColumn、Table.Combine、Table.UnpivotOtherColumns、Table.Pivot、Table.Group、Text.Split、Text.Combine、List.Sum、List.Distinct）。
- **逻辑经手工推演**：每段代码后面都附了预期输出表，数字与配套示例文件逐项一致（销售明细总计 42690、订单总计 21885）。

### 5.1 let/in 与 each（即 4.1 节代码的完整读法）

见 4.1 节。补充一点：`let` 里变量名用中文完全可以（如“提升标题”），M 代码里引用时写 `#"提升标题"` 这种带 `#""` 的形式是引用查询或步骤的标准写法，界面生成的代码就是这样。

### 5.2 拆分列与填充的 M 写法

```m
let
    源 = #"销售明细",
    拆分地址 = Table.SplitColumn(源, "地址", Splitter.SplitTextByDelimiter("-"), {"省", "市", "区"}),
    填充部门 = Table.FillDown(拆分地址, {"部门"}),
    合并省市 = Table.AddColumn(填充部门, "省市", each Text.Combine({[省], [市]}, "-"), type text)
in
    合并省市
```

读法：

- `Table.SplitColumn(表, 列名, 拆分器, 新列名列表)`：`Splitter.SplitTextByDelimiter("-")` 是按分隔符拆分的拆分器函数。
- `Table.FillDown(表, {"部门"})`：把“部门”列的空单元格用上方的值向下填充。
- `Text.Combine({[省], [市]}, "-")`：把两列文本用 `-` 连起来，得到“广东-深圳”。也可以写 `[省] & "-" & [市]`，`&` 是 M 的文本连接运算符。

预期输出（以地址“广东-深圳-南山区”、部门首行“华东”其后两行空为例）：拆出（省/市/区）三列，部门列三行全为（华东），新增（省市）列为（广东-深圳）。

### 5.3 从文件夹合并的 M 写法

```m
let
    源 = Folder.Files("D:\月报"),
    只留CSV = Table.SelectRows(源, each [Extension] = ".csv"),
    加内容 = Table.AddColumn(只留CSV, "内容", each Csv.Document([Content], [Delimiter=",", Encoding=65001])),
    展开内容 = Table.ExpandTableColumn(加内容, "内容", {"日期", "销售员", "地区", "产品", "数量", "单价", "金额"}),
    删多余列 = Table.RemoveColumns(展开内容, {"Content", "Name", "Extension", "Date accessed", "Date modified", "Date created", "Attributes", "Folder Path"})
in
    删多余列
```

读法：

- `Folder.Files("D:\月报")` 返回一张表，每行一个文件，含（Name）、（Extension）、（Content）等列，其中（Content）是文件内容的二进制值。
- `each [Extension] = ".csv"` 只留下 csv 文件；`Csv.Document([Content], ...)` 把每行的二进制内容解析成一张小表。
- `Table.ExpandTableColumn` 把嵌套的小表展开成普通列；最后删掉文件夹自带的多余列。

预期输出（以配套 3 个 csv 为例）：12 行、7 列（日期、销售员、地区、产品、数量、单价、金额），总计 42690，与“销售明细”表逐项一致。界面点“合并文件”生成的代码比这段更长（会多一个转换文件的函数），但逻辑相同；这段是可读的手写简化版。

### 5.4 合并查询的 M 写法

```m
let
    订单 = #"订单",
    客户档案 = #"客户档案",
    合并 = Table.NestedJoin(订单, {"客户编号"}, 客户档案, {"客户编号"}, "档案", JoinKind.LeftOuter),
    展开 = Table.ExpandTableColumn(合并, "档案", {"客户名称", "客户等级"}, {"客户名称", "客户等级"})
in
    展开
```

读法：

- `Table.NestedJoin(左表, 左键列表, 右表, 右键列表, 新列名, 联接类型)`：界面“合并查询”生成的就是这个函数。键用列表表示，多列关联就写多个列名，如 `{"省", "市"}`。
- `JoinKind.LeftOuter` 是默认值，不写也是左外部；换成 `JoinKind.Inner`、`JoinKind.LeftAnti`、`JoinKind.FullOuter`、`JoinKind.RightOuter`、`JoinKind.RightAnti` 即对应 03-合并与追加篇的六种联接。
- `Table.ExpandTableColumn` 把嵌套的档案表展开，只取“客户名称”与“客户等级”两列。

预期输出（左外部）：6 行，O001 至 O006 各带客户名称；O004（K999）客户名称与客户等级为空。把 `JoinKind.LeftOuter` 换成 `JoinKind.Inner` 得到 5 行（合计 17385）；换成 `JoinKind.LeftAnti` 得到 1 行（O004）。

### 5.5 逆透视与透视的 M 写法

```m
// 逆透视：把（月度交叉表）的 3 个月份列拆长
逆透视 = Table.UnpivotOtherColumns(#"月度交叉表", {"姓名"}, "月份", "销售额")

// 透视：把明细按（姓名、月份）转回交叉表（先按姓名和月份求和，再透视）
let
    源 = #"销售明细",
    按月汇总 = Table.Group(源, {"销售员", "月份"}, {{"月金额", each List.Sum([金额]), type number}}),
    透视 = Table.Pivot(按月汇总, List.Distinct(按月汇总[月份]), "月份", "月金额", List.Sum)
in
    透视
```

读法：

- `Table.UnpivotOtherColumns(表, 保留列, 属性列名, 值列名)`：除“姓名”外的列全部逆透视，新列名（月份）与（销售额）。新增月份列出现时自动跟进，比“仅逆透视选中列”更适合月报。
- `Table.Group(表, 分组列, 聚合列表)`：`{{"月金额", each List.Sum([金额]), type number}}` 表示按组求和，`each` 里的 `[金额]` 是当前组的金额列。
- `Table.Pivot(表, 透视值列表, 属性列, 值列, 聚合函数)`：`List.Distinct(按月汇总[月份])` 取出所有不重复的月份作为新列名。

预期输出：逆透视得到 9 行（见 02-转换操作篇的表，总计 42690）；透视得到 3 行 4 列的交叉表（姓名、1月、2月、3月），与配套“月度交叉表”一致。

### 5.6 自定义函数：把文件路径变成参数

```m
// 查询 fn读销售CSV：输入文件路径，输出清洗好的表
(路径 as text) as table =>
let
    源 = Csv.Document(File.Contents(路径), [Delimiter=",", Encoding=65001]),
    提升标题 = Table.PromoteHeaders(源),
    改类型 = Table.TransformColumnTypes(提升标题, {{"日期", type date}, {"金额", type number}}),
    筛选 = Table.SelectRows(改类型, each [金额] > 0)
in
    筛选
```

读法：`(路径 as text) as table => ...` 定义一个函数：输入文本型路径，输出表。调用时写 `fn读销售CSV("D:\月报\销售2026-01.csv")`。

要点：函数查询要在编辑器里右键查询→“创建函数”或直接在高级编辑器里改写；函数本身不返回数据，只在被调用时才执行；路径写死在代码里是新手最常见的坑，参数化之后换个文件夹只需改调用处。

### 5.7 从数据库与网页取数的 M 写法

```m
// SQL Server：只取 2026 年的订单
源 = Sql.Database("服务器名", "销售库", [Query="select * from 订单 where year(日期) = 2026"])

// 工作簿：读指定工作表（Excel.Workbook 返回导航表，需再展开要用的表）
源 = Excel.Workbook(File.Contents("D:\数据\产品目录.xlsx"))

// 网页 JSON 接口：先取内容，再解析为 JSON（把地址换成真实接口地址）
源 = Json.Document(Web.Contents("https://api.example.com/sales"))
```

读法：`Sql.Database` 的第 3 个参数是选项记录，`[Query=...]` 可以直接贴 SQL；`Excel.Workbook` 读出的是导航表，要在导航器里点开目标工作表才会展开；`Web.Contents` 取回二进制内容，`Json.Document` 解析成记录或列表。

要点：这三段是连接模式的写法，服务器名、文件路径、接口地址都要换成真实值；网页接口需要账号密码时，`Web.Contents` 要加凭据选项，或在界面里按提示输入。

## 6. 实际工作场景

### 6.1 批量改类型：界面点 10 次，代码改一行

10 个月度文件合并后，金额列全是文本型。界面做法是逐个查询点类型图标；M 做法是在高级编辑器里找到 `Table.TransformColumnTypes` 那行，把 `type text` 改成 `type number`，一次全改。改完刷新，10 个查询一起变。

### 6.2 自定义函数复用：12 个月共用一个清洗函数

12 个月度文件结构相同、清洗步骤相同。把清洗逻辑写成一个函数查询 `fn读销售CSV`（见 5.6 节），12 个查询都调它：换个文件夹只需改调用处的路径，清洗逻辑改一处、12 处全生效。函数本身不返回数据，只在被调用时才执行。

## 7. 常见错误

| 错误现象 | 原因 | 解决方案 |
|---|---|---|
| 公式栏标红：找不到函数 | M 区分大小写，函数名写错了大小写 | 核对函数名：Table.SelectRows 不能写成 table.selectrows；函数名与参数签名对照 Microsoft Learn 核对 |
| 文本里的双引号报错 | 文本用双引号，文本里的双引号要写两遍 | 文本里的一个双引号写成两个，如 `"他说""你好"""` |
| 取列表元素差一位 | 列表从 0 开始数，{1} 是第 2 个元素 | Text.Split 返回的列表从 0 数：第 1 个元素写 {0}，第 2 个写 {1} |
| 高级编辑器改完查询变空 | let/in 里最后 in 后面的输出变量名写错了 | 检查 in 后面是不是最后一步的变量名；变量名改了，引用处要跟着改 |
| 自定义函数不返回数据 | 函数本身不执行，只在被调用时才执行 | 新建一个查询调用它，如 `fn读销售CSV("D:\月报\销售2026-01.csv")`，再看结果 |
| 改了列名，后面步骤全红 | 后面的步骤还在引用旧列名 | 点报错的步骤，看公式栏里的旧列名，逐个改成新列名；或删掉报错步骤重做 |

## 8. 相关教程

- 同栏目：获取数据（docs/powerquery/01-power-query.md，Csv.Document、Folder.Files、Sql.Database 的来源）
- 同栏目：转换操作（docs/powerquery/02-transform.md，拆分、填充、逆透视的界面做法）
- 同栏目：合并与追加（docs/powerquery/03-merge-append.md，NestedJoin 与 Combine 的界面做法）
- 同栏目：查询管理与刷新（docs/powerquery/05-query-management.md，改代码后刷新与错误处理）
- 跨栏目：数据透视表（docs/pivottable/01-pivottable.md，M 清洗后的表直接做透视表）

## 9. 练习题

1. “地址”列的值是“广东-深圳-南山区”，要取“深圳”。写出 M 代码，并说出 Text.Split 返回的是什么结构、取第 2 个元素为什么写 {1}。
2. `let...in` 结构里，`in` 后面应该写什么？如果写错了会怎样？
3. 把文件路径变成函数参数有什么好处？函数本身会返回数据吗？
4. M 语言区分大小写吗？`table.selectrows` 能用吗？
5. 界面点“合并文件”生成的代码比 5.3 节的手写版更长，为什么？逻辑相同吗？

## 10. 参考答案

<details>
<summary>点击查看答案</summary>

1. `Table.AddColumn(上一步, "城市", each Text.Split([地址], "-"){1})`。Text.Split 返回列表 `{"广东", "深圳", "南山区"}`，列表从 0 开始数，{1} 是第 2 个元素（深圳）。

2. `in` 后面写最后一步的变量名（即输出）。写错了查询会变空或报错，要检查 in 后面是不是最后一步的变量名。

3. 换个文件夹只需改调用处的路径，清洗逻辑改一处、多处全生效。函数本身不返回数据，只在被调用时才执行。

4. 区分大小写。`table.selectrows` 不能用，必须写 `Table.SelectRows`。

5. 界面版多一个转换文件的函数（处理示例文件与类型推断），手写版是可读的简化版；逻辑相同，都是读文件夹、筛 csv、解析、展开、删多余列。

</details>

## 11. 可下载示例文件

- [PowerQuery-示例-销售明细.xlsx](../../templates/PowerQuery-示例-销售明细.xlsx)（销售明细 12 行、订单 6 行、客户档案 5 行、月度交叉表 3 行、使用说明；总计 42690）
- [销售2026-01.csv](../../templates/销售2026-01.csv)（1 月 4 行，14690）
- [销售2026-02.csv](../../templates/销售2026-02.csv)（2 月 4 行，9730）
- [销售2026-03.csv](../../templates/销售2026-03.csv)（3 月 4 行，18270）
