# Excel 学习路线图：4 条路线，从零基础到自动化

- **所属栏目**：C18 学习路线与练习考试 ／ 学习路线图
- **适用版本**：Excel 2016／2019／2021／365
- **难度**：入门到高级（4 条路线按难度递进）
- **预计阅读时长**：30 分钟（选路与规划；按路线学完约 40–80 小时）
- **更新日期**：2026-11-10

## 1. 学习目标

学完本篇，你将能够：

- 根据自己的目标（日常办公、进阶提效、数据分析、自动化）选对一条路线，不贪多、不跳阶段
- 按阶段顺序学完所选路线的全部教程，每阶段做完检查点再进入下一阶段
- 分清两类教程的打开方式：有 md 文件的点链接直达；只有站内数据的按 id 在站内搜索
- 学完一条路线后，独立完成该路线出口处列出的工作任务

## 2. 基础概念

4 条路线总览如下。路线可以交叉：一篇教程可以出现在多条路线里，学过的不必重学；4 条路线合起来覆盖本站全部 92 篇教程，没有一篇被落下。

| 路线 | 适合人群 | 前置要求 | 预计学时 | 出口能力 |
| --- | --- | --- | --- | --- |
| 零基础路线 | 从没系统学过 Excel 的新手 | 无，会开关电脑即可 | 约 40 小时 | 独立做格式规范的表格、基础公式统计、透视汇总与简单图表 |
| 进阶路线 | 会基础操作，想把表格做得又快又规范的办公人员 | 已掌握零基础路线第 1–4 阶段 | 约 25 小时 | 规范输入、条件格式、图表选型与透视表进阶分析 |
| 数据分析路线 | 要做报表、看板与多表分析的数据岗与业务岗 | 已掌握公式与透视表基础 | 约 35 小时 | Power Query 清洗、数据模型与 DAX 度量值、商业场景分析 |
| 自动化路线 | 要把重复工作交给程序的进阶用户 | 已掌握公式与 Power Query 基础 | 约 30 小时 | VBA 进阶、Office Scripts、Python 批量处理、自动出报表 |

补充说明：

- 函数百科（C04，100 个函数）是查阅型内容，不计入 92 篇教程；学到公式时按需查阅，不必先背完再动手。
- VBA 基础语法、对象模型、关键字手册与实例精讲在 C11 VBA 栏目中，本路线所列的 VBA 教程为进阶专篇（id: vba-05-advanced）。
- 站内数据教程没有独立的 md 文件，本篇用纯文字写出标题与站内 id，不放超链接；有 md 文件的教程一律用相对路径链接，可直接点开。

## 3. 操作步骤

1. 定目标：先想清楚学完要解决什么工作问题，再对照第 2 节的总览表选一条主路线。
2. 选路线：一次只走一条主路线；需要另一条路线的内容时，按本篇给出的阶段顺序补学，不必从头重走。
3. 按阶段学：每个阶段 3–5 篇教程，按编号顺序学；每篇教程跟着做一遍，不要只看。
4. 做练习：学完一个阶段就做 [练习题库：30 道题，跟着真实数据练](./01-exercises.md)（id: practice-01-exercises）里对应的题目。
5. 自测过关：对照每阶段末尾的检查点自测，通过后再进下一阶段；没通过就重做该阶段的教程，不要带着缺口往下走。

## 4. 零基础路线

- **适合人群**：从没系统学过 Excel 的新手，日常要做表格、统计与简单汇报
- **前置要求**：无，会开关电脑、会打字即可
- **预计学时**：约 40 小时（6 个阶段，每阶段约 6–7 小时）
- **学完能做什么**：独立做出格式规范的表格；用公式做求和、计数与条件统计；用数据透视表做汇总；用图表汇报结果；完成销售、财务、人事、项目、营销 5 类基础办公表

### 4.1 第一阶段：认识 Excel 与数据录入

1. [认识 Excel：界面、工作簿与工作表](../basics/01-getting-started.md)（id: basics-01-getting-started）
2. [数据录入与编辑：输入技巧、填充柄与复制粘贴](../basics/02-data-entry.md)（id: basics-02-data-entry）
3. [单元格与区域操作：选定、行列、合并与名称框](../basics/03-cells-ranges.md)（id: basics-03-cells-ranges）
4. 认识 Excel 界面（站内数据，id: interface）
5. 数据录入技巧（站内数据，id: enter）

检查点：不看教程，能说出功能区、编辑栏、名称框的位置；能在 10 分钟内录入一张 20 行的表格并用填充柄补齐序号。

### 4.2 第二阶段：格式与打印

1. [数字格式：常规、货币、日期、百分比与自定义格式代码](../formatting/01-number-format.md)（id: formatting-01-number-format）
2. [字体与对齐：字号颜色、对齐缩进与自动换行](../formatting/02-font-alignment.md)（id: formatting-02-font-alignment）
3. [边框与填充：边框样式、填充图案与网格线](../formatting/03-border-fill.md)（id: formatting-03-border-fill）
4. 单元格格式（站内数据，id: format）
5. 打印设置（站内数据，id: print）

检查点：能把一张原始数据表排成标题、表头、金额列、日期列格式都正确的表格；打印预览时每页都有表头、无内容被截断。

### 4.3 第三阶段：数据整理入门

1. [排序与筛选：单条件、多条件、自动筛选与高级筛选](../data-management/01-sort-filter.md)（id: data-management-01-sort-filter）
2. [数据清洗：分列、去重、TRIM/CLEAN 与查找替换清理](../data-management/02-clean-data.md)（id: data-management-02-clean-data）
3. 排序与筛选（站内数据，id: sortfilter）
4. [超级表：创建、结构化引用、表样式与汇总行](../formatting/05-tables.md)（id: formatting-05-tables）

检查点：能对一份带重复值与空格的数据做去重与清洗；能按两个条件排序并筛选出目标行；能把数据区域转成超级表并加汇总行。

### 4.4 第四阶段：公式入门

1. [公式入门：公式结构、运算符与输入编辑](../formulas/01-formula-basics.md)（id: formulas-01-formula-basics）
2. [引用方式：相对引用、绝对引用、混合引用与跨表引用](../formulas/02-references.md)（id: formulas-02-references）
3. [名称与定义：定义名称、名称管理器与动态名称](../formulas/03-names.md)（id: formulas-03-names）
4. [公式审核与调试：追踪引用、公式求值与错误检查](../formulas/05-auditing.md)（id: formulas-05-auditing）

检查点：能说出相对引用与绝对引用的区别，并用 $ 锁定一行公式下拉不出错；公式报错时能用追踪引用找到源头。

### 4.5 第五阶段：汇总与展示入门

1. [数据透视表](../pivottable/01-pivottable.md)（id: pivottable）
2. [条件格式](../conditional-formatting/01-conditional-formatting.md)（id: conditional-formatting）
3. [柱形图与条形图](../charts/01-column-bar.md)（id: charts-01-column-bar）
4. 图表类型选择指南（站内数据，id: choose）
5. 图表制作步骤（站内数据，id: make）

检查点：能用数据透视表对销售明细按月、按人汇总；能用条件格式标出超标的数值；能按数据特点选对图表类型并做出一张可汇报的图。

### 4.6 第六阶段：办公实战入门

1. [销售数据分析：从 12 行订单到销售仪表盘](../business/01-sales-analysis.md)（id: business-01-sales-analysis）
2. [财务报表制作：利润表、资产负债表、现金流量表](../business/02-financial-statements.md)（id: business-02-financial-statements）
3. [人力资源管理：考勤统计、工资条与花名册](../business/03-hr-management.md)（id: business-03-hr-management）
4. [项目管理：项目计划表、甘特图与进度跟踪](../business/04-project-management.md)（id: business-04-project-management）
5. [市场营销分析：营销漏斗、渠道对比与活动复盘](../business/05-marketing-analysis.md)（id: business-05-marketing-analysis）

检查点：能任选其中两类业务场景，从原始数据出发独立做完一张可交付的分析表；做完后能向别人讲清每一步公式的含义。

## 5. 进阶路线

- **适合人群**：会基础操作，想把表格做得又快又规范的办公人员
- **前置要求**：已掌握零基础路线第 1–4 阶段（录入、格式、整理、公式入门）
- **预计学时**：约 25 小时（6 个阶段，每阶段约 4 小时）
- **学完能做什么**：做出防错的输入模板；用条件格式让异常一眼可见；按数据特点选对图表并美化；用透视表进阶功能做多维分析

### 5.1 第一阶段：公式进阶

1. [高级公式技巧：LET、LAMBDA、嵌套与条件求和进阶](../formulas/06-advanced-tricks.md)（id: formulas-06-advanced-tricks）
2. [数组公式](../formulas/04-array-formulas.md)（id: array-formulas）
3. [超级表应用（交叉引用索引）](../data-management/03-tables-ref.md)（id: data-management-03-tables-ref）
4. [样式与主题：单元格样式、套用表格格式与主题配色](../formatting/04-styles-themes.md)（id: formatting-04-styles-themes）

检查点：能用 LET 把一串嵌套公式拆成可读的步骤；能说出动态数组溢出的含义；能用结构化引用写出可读性更好的公式。

### 5.2 第二阶段：数据验证与规范输入

1. [数据验证](../data-management/04-data-validation.md)（id: data-validation）
2. 下拉菜单（站内数据，id: dval-0）
3. 联动下拉框（站内数据，id: dval-1）
4. 日期限制（站内数据，id: dval-2）
5. 输入规则（站内数据，id: dval-3）

检查点：能做出带下拉菜单与日期限制的输入模板；输入非法值时能被拦截并看到提示；能做出省市联动的二级下拉。

### 5.3 第三阶段：条件格式进阶

1. 颜色标记（站内数据，id: cfmt-0）
2. 数据条（站内数据，id: cfmt-1）
3. 图标集（站内数据，id: cfmt-2）
4. 自定义公式（站内数据，id: cfmt-3）
5. 案例：自动标红逾期（站内数据，id: cfmt-4）

检查点：能用数据条与图标集做直观对比；能写公式规则实现整行高亮；能复现"自动标红逾期"的案例。

### 5.4 第四阶段：图表进阶

1. [折线图与面积图](../charts/02-line-area.md)（id: charts-02-line-area）
2. [饼图与圆环图](../charts/03-pie-doughnut.md)（id: charts-03-pie-doughnut）
3. [散点图与气泡图](../charts/04-scatter-bubble.md)（id: charts-04-scatter-bubble）
4. [组合图](../charts/05-combo.md)（id: charts-05-combo）
5. 图表美化技巧（站内数据，id: beauty）

检查点：能按"比较、趋势、占比、关系"四类场景选对图表；能做出双坐标轴组合图；能按美化要点删繁就简，标题写成结论句。

### 5.5 第五阶段：透视表进阶

1. 基础操作（站内数据，id: pivot_adv-0）
2. 分组统计（站内数据，id: pivot_adv-1）
3. 多维分析（站内数据，id: pivot_adv-2）
4. 切片器（站内数据，id: pivot_adv-3）
5. 时间轴（站内数据，id: pivot_adv-4）

检查点：能对日期字段做按月、按季度分组；能用切片器与时间轴做交互筛选；能用值显示方式做出占比与累计。

### 5.6 第六阶段：图表收尾与综合

1. [新图表类型](../charts/06-new-charts.md)（id: charts-06-new-charts）
2. [图表格式化](../charts/07-formatting.md)（id: charts-07-formatting）
3. 案例：高亮重复（站内数据，id: cfmt-5）
4. 圈释无效数据（站内数据，id: dval-4）
5. [数据导入导出：文本/CSV 导入、从数据库与网页获取、导出格式](../data-management/05-import-export.md)（id: data-management-05-import-export）

检查点：能说出瀑布图、漏斗图、直方图的适用场景；能用圈释无效数据找出模板里已有的非法输入；能从 CSV 导入数据并正确处理编码与分隔符。

## 6. 数据分析路线

- **适合人群**：要做报表、看板与多表分析的数据岗与业务岗
- **前置要求**：已掌握公式与透视表基础（零基础路线第 4–5 阶段，或同等水平）
- **预计学时**：约 35 小时（6 个阶段，每阶段约 6 小时）
- **学完能做什么**：用 Power Query 把每月重复清洗做成一键刷新；建数据模型写 DAX 度量值；做出 KPI 看板与经营仪表盘；独立完成财务、销售、人事、库存 4 类分析场景

### 6.1 第一阶段：数据获取与清洗

1. [Power Query](../powerquery/01-power-query.md)（id: powerquery）
2. 导入 CSV（站内数据，id: pq_adv-0）
3. 导入网页（站内数据，id: pq_adv-1）
4. 数据清洗（站内数据，id: pq_adv-2）

检查点：能从 CSV 与网页取数；能完成拆分列、填充、去重等清洗步骤；能说出"应用的步骤"是什么，改一步后刷新结果正确。

### 6.2 第二阶段：Power Query 进阶

1. 合并表（站内数据，id: pq_adv-3）
2. 拆分列（站内数据，id: pq_adv-4）
3. 自动刷新（站内数据，id: pq_adv-5）
4. [数据导入导出：文本/CSV 导入、从数据库与网页获取、导出格式](../data-management/05-import-export.md)（id: data-management-05-import-export）

检查点：能合并两个结构不同的表并处理联接类型；换一份同结构的新源文件后，点刷新就能得到正确结果。

### 6.3 第三阶段：透视与动态图表

1. 数据模型（站内数据，id: pivot_adv-5）
2. 动态图表（站内数据，id: dynamic）
3. [练习题库：30 道题，跟着真实数据练](./01-exercises.md)（id: practice-01-exercises）
4. [仓库与物流：库存管理、出入库流水与物流单据](../business/06-inventory-logistics.md)（id: business-06-inventory-logistics）

检查点：能用透视表的数据模型功能关联多张表；能做出下拉菜单联动的动态图表；能独立完成练习题库里的透视表题目。

### 6.4 第四阶段：数据模型与 DAX 基础

1. [数据模型：把多张表变成一个可分析的整体](../powerpivot-dax/01-data-model.md)（id: powerpivot-dax-01-data-model）
2. [DAX 基础：语法、计算列与度量值、行上下文与筛选上下文](../powerpivot-dax/02-dax-basics.md)（id: powerpivot-dax-02-dax-basics）
3. 数据模型（站内数据，id: ppivot-0）
4. 多表关联（站内数据，id: ppivot-1）
5. DAX 公式入门（站内数据，id: ppivot-2）

检查点：能说出计算列与度量值的区别；能建星型模型并正确设置表间关系；能写出第一个可用的度量值。

### 6.5 第五阶段：DAX 进阶与 KPI

1. [度量值与计算列：CALCULATE、SUMX、RELATED、FILTER、ALL 与 KPI](../powerpivot-dax/03-measures-calculated-columns.md)（id: powerpivot-dax-03-measures-calculated-columns）
2. [时间智能：日期表、DATESYTD、TOTALYTD、SAMEPERIODLASTYEAR、DATEADD 与同比环比](../powerpivot-dax/04-time-intelligence.md)（id: powerpivot-dax-04-time-intelligence）
3. [高级 DAX 与商业案例：排名、TOPN、动态度量值、销售仪表盘与库存周转](../powerpivot-dax/05-advanced-dax-cases.md)（id: powerpivot-dax-05-advanced-dax-cases）
4. 常用 DAX（站内数据，id: ppivot-3）
5. KPI 指标（站内数据，id: ppivot-4）

检查点：能用 CALCULATE 改写筛选上下文；能做出同比与环比度量值；能定义 3 个以上 KPI 度量值并放进透视表验证结果。

### 6.6 第六阶段：商业场景实战

1. 财务报销统计（站内数据，id: scenario-0）
2. 销售业绩分析（站内数据，id: scenario-1）
3. 人事考勤统计（站内数据，id: scenario-2）
4. 库存预警管理（站内数据，id: scenario-3）
5. 经营仪表盘（站内数据，id: scenario-4）

检查点：能任选其中两个场景，从原始数据出发独立复现完整分析过程；能把结果拼成一页经营仪表盘并讲清每个 KPI 的口径。

## 7. 自动化路线

- **适合人群**：要把重复工作交给程序的进阶用户，日常有固定报表与批量处理任务
- **前置要求**：已掌握公式与 Power Query 基础（零基础路线第 4 阶段与数据分析路线第 1 阶段，或同等水平）
- **预计学时**：约 30 小时（5 个阶段，每阶段约 6 小时）
- **学完能做什么**：用 VBA 或 Office Scripts 自动生成报表；用 Python 批量处理几十个工作簿；把月度报表做成"换源文件、点一下就出结果"的半自动流程

### 7.1 第一阶段：自动化入门与版本

1. [版本与兼容：从 Excel 2003 到 Microsoft 365](../basics/05-versions.md)（id: basics-05-versions）
2. [高级公式技巧：LET、LAMBDA、嵌套与条件求和进阶](../formulas/06-advanced-tricks.md)（id: formulas-06-advanced-tricks）
3. [Power Query](../powerquery/01-power-query.md)（id: powerquery）

检查点：能说出自己用的 Excel 版本支持哪些自动化方式（VBA、Office Scripts、Python in Excel 各有什么版本要求）；能用 LET 与 LAMBDA 把重复计算封装成可复用的公式。

### 7.2 第二阶段：VBA 进阶

1. [VBA 进阶：用户窗体、类模块、加载宏与错误处理进阶](../vba/05-advanced.md)（id: vba-05-advanced）
2. [销售数据分析：从 12 行订单到销售仪表盘](../business/01-sales-analysis.md)（id: business-01-sales-analysis）
3. [数据清洗：分列、去重、TRIM/CLEAN 与查找替换清理](../data-management/02-clean-data.md)（id: data-management-02-clean-data）

检查点：能读懂一段带用户窗体的 VBA 代码并说出每部分的作用；能把"清洗加汇总"的手工流程拆成可用宏录制的步骤清单；VBA 基础语法、对象模型与关键字见 C11 VBA 栏目。

### 7.3 第三阶段：Office Scripts

1. [Office Scripts 入门：不用 VBA 也能让 Excel 自己干活](../office-scripts/01-introduction.md)（id: office-scripts-01-introduction）
2. [TypeScript 基础：写出你的第一段 Office Scripts 代码](../office-scripts/02-typescript-basics.md)（id: office-scripts-02-typescript-basics）
3. [自动化实战：定时任务、Power Automate 联动与 3 个完整案例](../office-scripts/03-automation-cases.md)（id: office-scripts-03-automation-cases）
4. [市场营销分析：营销漏斗、渠道对比与活动复盘](../business/05-marketing-analysis.md)（id: business-05-marketing-analysis）

检查点：能说出 Office Scripts 与 VBA 的适用场景区别；能录制并改写一段脚本；能把一个脚本挂到 Power Automate 做定时执行。

### 7.4 第四阶段：Python 与 Excel

1. [openpyxl：用 Python 读写 Excel 工作簿](../python-excel/01-openpyxl.md)（id: python-excel-01-openpyxl）
2. [pandas：DataFrame 与 Excel 互转、数据清洗与分组汇总](../python-excel/02-pandas.md)（id: python-excel-02-pandas）
3. [xlwings：调用正在运行的 Excel，批量处理工作簿](../python-excel/03-xlwings.md)（id: python-excel-03-xlwings）
4. [Python in Excel：用 =PY() 在单元格里写 Python](../python-excel/04-python-in-excel.md)（id: python-excel-04-python-in-excel）
5. [数据导入导出：文本/CSV 导入、从数据库与网页获取、导出格式](../data-management/05-import-export.md)（id: data-management-05-import-export）

检查点：能用 openpyxl 读写一个工作簿；能用 pandas 对一份数据做分组汇总并写回 Excel；能说出 openpyxl、xlwings、Python in Excel 各自适合什么任务。

### 7.5 第五阶段：综合自动化项目

1. 经营仪表盘（站内数据，id: scenario-4）
2. [财务报表制作：利润表、资产负债表、现金流量表](../business/02-financial-statements.md)（id: business-02-financial-statements）
3. [练习题库：30 道题，跟着真实数据练](./01-exercises.md)（id: practice-01-exercises）
4. [高级 DAX 与商业案例：排名、TOPN、动态度量值、销售仪表盘与库存周转](../powerpivot-dax/05-advanced-dax-cases.md)（id: powerpivot-dax-05-advanced-dax-cases）

检查点：能独立设计一个"换源文件、点一下就出结果"的月度报表流程，并写清每一步用的是公式、Power Query、VBA 还是 Python；能完整做完练习题库里对应的综合题。

## 8. 练习题

自测清单（做完对应阶段后回答，不必翻教程先想一遍）：

1. 学完零基础路线第 4 阶段，你能说出相对引用与绝对引用的区别吗？把 =A1*$B$1 从 C1 复制到 C2，公式会变成什么？
2. 学完进阶路线第 2 阶段，你要给同事做一张"只能从下拉里选部门、日期不能超过今天"的录入表，用哪几篇教程里的方法？
3. 学完进阶路线第 5 阶段，一份 5000 行销售明细，老板要"每个销售员每个月的销售额，还能按季度筛选"，你第一步做什么？
4. 学完数据分析路线第 2 阶段，下个月拿到一份同结构的新源文件，你要做哪几步才能让上个月的清洗结果自动重算？
5. 学完自动化路线第 3 阶段，Office Scripts 与 VBA 各适合什么场景？你会怎么选？
6. 学完数据分析路线第 5 阶段，计算列与度量值有什么区别？同比增长率应该写成哪一种？

## 9. 参考答案

1. 相对引用会随公式位置移动，绝对引用（带 $）固定不变；=A1*$B$1 从 C1 复制到 C2 后变成 =A2*$B$1（A1 变 A2，$B$1 不变）。对应教程：[引用方式：相对引用、绝对引用、混合引用与跨表引用](../formulas/02-references.md)（id: formulas-02-references）。
2. 用数据验证做下拉菜单与日期限制：先看 [数据验证](../data-management/04-data-validation.md)（id: data-validation），再按"下拉菜单（站内数据，id: dval-0）"与"日期限制（站内数据，id: dval-2）"操作；部门联动再加"联动下拉框（站内数据，id: dval-1）"。
3. 第一步先确认源数据是规范的二维表（首行标题、无合并单元格、无空行空列），再插入数据透视表，行放销售员、列放月份、值放销售额求和，季度筛选可用分组或切片器。对应教程：[数据透视表](../pivottable/01-pivottable.md)（id: pivottable）与"分组统计（站内数据，id: pivot_adv-1）"、"切片器（站内数据，id: pivot_adv-3）"。
4. 只要源文件结构不变：把新文件放到原路径或重新指定源，然后点"刷新"，Power Query 会按"应用的步骤"自动重跑全部清洗步骤，不必重做。对应教程：[Power Query](../powerquery/01-power-query.md)（id: powerquery）与"自动刷新（站内数据，id: pq_adv-5）"。
5. VBA 适合桌面版 Excel 的本地自动化（用户窗体、加载宏、复杂交互）；Office Scripts 适合网页版 Excel 与 Power Automate 定时任务。按"用哪一端 Excel"和"要不要定时跑"来选。对应教程：[Office Scripts 入门：不用 VBA 也能让 Excel 自己干活](../office-scripts/01-introduction.md)（id: office-scripts-01-introduction）与 [VBA 进阶：用户窗体、类模块、加载宏与错误处理进阶](../vba/05-advanced.md)（id: vba-05-advanced）。
6. 计算列逐行算、存在表里；度量值按当前筛选上下文动态算、不存数据。同比增长率是聚合后比较，应该写成度量值。对应教程：[DAX 基础：语法、计算列与度量值、行上下文与筛选上下文](../powerpivot-dax/02-dax-basics.md)（id: powerpivot-dax-02-dax-basics）与"DAX 公式入门（站内数据，id: ppivot-2）"。
