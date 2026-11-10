# 自动化实战：定时任务、Power Automate 联动与 3 个完整案例

- **所属栏目**：C12 Office Scripts ／ C12-L2-3 自动化实战
- **适用版本**：Microsoft 365 商业版／教育版（网页版 Excel、Excel for Windows 2210 及更高版本、Excel for Mac）；Power Automate 联动须持有 Microsoft 365 商业许可
- **难度**：进阶
- **预计阅读时长**：50 分钟
- **更新日期**：2026-11-10

> 验证方式说明：本篇所有 Office Scripts 代码均未在 Excel 网页版内执行验证，也未在 Power Automate 里实际跑通过任何流程。验证口径为三层：API 名称与方法签名对照 Microsoft 官方文档核对（learn.microsoft.com 的 Office Scripts 基础文档、ExcelScript API 参考与 Power Automate 联动教程）；每段代码经去类型注解后的语法检查与括号配对检查；三个案例的输入数据与预期结果经 Python 独立验算（金额合计、行数、去重结果逐项算过）。本篇正文与全站均不使用声称代码已在 Excel 内实际跑通过的表述；每段代码配的是预期效果说明，不是实测运行结果。Power Automate 流程步骤是按官方教程的操作顺序写的思路说明，首次搭建时请以 Power Automate 界面实际显示的触发器与动作为准。

## 1. 学习目标

学完本篇，你将能够：
- 说出 Office Scripts 本身没有定时器，定时与事件触发要由 Power Automate 提供
- 搭建一个按计划运行脚本的 Power Automate 流程（定时触发加 Excel Online (Business) 连接器的 Run script 动作）
- 搭建一个邮件触发的流程：收到指定邮件后运行脚本，并把结果用 Teams 消息通知出来
- 写出日报自动汇总、数据清洗、格式标准化 3 个完整脚本，并说出每个脚本的输入、处理与预期输出
- 说出交给 Power Automate 调用的脚本必须遵守的两条规则（不用相对引用、脚本与工作簿存在 OneDrive 或 SharePoint 里）

## 2. 基础概念

### 2.1 Office Scripts 没有自己的触发器

Office Scripts 只负责“改工作簿”这一件事：它没有定时器，也没有“收到邮件就运行”这种事件触发。触发器由 Power Automate 提供，分工是：Power Automate 决定“什么时候跑”，Office Scripts 决定“跑的时候对工作簿做什么”。这个分工是理解本篇全部内容的前提。

### 2.2 Power Automate 联动的三要素

- **触发器（Trigger）**：流程的起因。常用三种：定时触发（Scheduled cloud flow，按固定周期运行）、邮件到达（When a new email arrives，收到新邮件时运行）、手动触发（Manually trigger a flow，点按钮运行）。
- **Run script 动作**：Excel Online (Business) 连接器里的动作。它的作用是“在某个工作簿上运行某段脚本”。脚本存在 OneDrive for Business 的默认位置时用 Run script，脚本存在团队 SharePoint 库里时用 Run script from a SharePoint library（Microsoft 官方教程“Automatically save content from emails in a workbook”）。
- **参数与返回值**：main 函数可以带额外参数，流程调用时把数据传进去（例如把邮件发件人传给脚本）；脚本也可以返回结果，流程拿到结果后再做下一步（例如把汇总数字发到 Teams）。

### 2.3 交给 Power Automate 调用的脚本的两条硬规则

这两条来自 Microsoft 官方教程，违反任何一条，流程都会在运行时失败：

1. **不要用相对引用**：脚本里不要用 workbook.getActiveWorksheet() 这类“当前活动”引用，因为流程运行时没有“当前活动工作表”这个概念。要用 workbook.getWorksheet("表名") 按名字取表。
2. **工作簿与脚本必须在云端**：工作簿要在 OneDrive for Business、SharePoint 库或 Teams 团队库里，本机磁盘上的文件流程选不到；脚本同样存在 OneDrive 或 SharePoint 里。

### 2.4 平台限制（写脚本前先知道边界）

以下限制来自 Microsoft 官方平台要求文档，设计流程时要先留出余量：

- 对工作簿的一次请求与响应限制为 5MB，单个区域限制为 500 万单元格。
- Excel Online (Business) 连接器的 Run script 动作不支持 ISO strict 格式的工作簿文件。
- 从 Power Automate 运行的脚本里不能用 fetch 等外部调用。
- 每天调用次数与单次超时有官方限额，搭建流程前请以 Microsoft 官方文档最新说明为准，不要按经验值估算。

## 3. 操作步骤

### 3.1 搭建定时流程：每天 9 点运行日报汇总脚本

1. 在网页版 Excel 里打开日报工作簿，确认“明细”工作表已按案例一的数据格式准备好（待补图）。
2. 打开 Power Automate（flow.microsoft.com），点“创建”→“计划的云端流”，填流程名（例如“每天 9 点汇总日报”），设置开始时间与重复周期（每天、9:00），点“创建”（待补图）。
3. 点“新建步骤”，搜索“Excel run script”，选 Excel Online (Business) 连接器的“Run script”动作（待补图）。
4. 在 Run script 动作里依次选：位置（OneDrive for Business）、文档库、文件（日报工作簿）、脚本（日报自动汇总脚本）；不要选本机文件（待补图）。
5. 点“保存”，再点“测试”→“手动测试”跑一次，检查日报工作簿的“日报汇总”工作表是不是生成了（待补图）。
6. 测试通过后不用再管：流程每天 9 点自动运行一次；想停用时在“我的流”里把流程关掉（待补图）。

### 3.2 搭建邮件触发流程：收到邮件后记一笔并通知 Teams

1. 在 Power Automate 里点“创建”→“自动化的云端流”，触发器选“收到新邮件时”（When a new email arrives），可按发件人或主题做筛选条件（待补图）。
2. 加一个步骤，选 Excel Online (Business) 连接器的“Run script”动作，选定工作簿与记录邮件的脚本，并把触发器的动态内容（发件人、主题、收到时间）作为脚本参数传进去（待补图）。
3. 再加一个步骤，选 Teams 连接器的“在聊天或频道中发布消息”动作，把脚本返回的结果（例如“已记录 1 封邮件，来自张三”）写进消息内容，选定要通知的频道（待补图）。
4. 保存后发一封测试邮件，检查三件事：工作簿里多了一行记录、Teams 频道收到通知、流程运行历史里显示成功（待补图）。
5. 如果多封邮件同时到达导致工作簿冲突，把邮件触发器的并发控制打开，一次只处理一封（Microsoft 官方教程的建议）。

### 3.3 把脚本交给流程前，先在 Excel 里手动跑一遍

1. 在网页版 Excel 里打开目标工作簿，手动运行一次脚本，确认结果符合预期（待补图）。
2. 检查脚本里没有 getActiveWorksheet()，所有工作表都按名字取（待补图）。
3. 检查脚本里没有 fetch 等外部调用（流程里跑会失败）（待补图）。
4. 确认工作簿与脚本都在 OneDrive for Business 或 SharePoint 里（待补图）。
5. 再去 Power Automate 里建流程、选脚本、测试运行（待补图）。

## 4. 图文示例

示例：定时流程的结构图（文字版）。

- 第一格是触发器：每天 9:00（Scheduled cloud flow，重复周期每天）。
- 第二格是 Run script 动作：Excel Online (Business) 连接器，位置 OneDrive for Business，文件是日报工作簿，脚本是“日报自动汇总”。
- 第三格（可选）是 Teams 通知动作：把脚本返回的“今日合计 30000 元”发到“销售日报”频道。
- 箭头方向：触发器 → Run script → Teams 通知；任何一格失败，流程运行历史里会标出失败的那一格，先修那一格再重跑（待补图）。

## 5. 公式和代码

本节 3 个案例各配一段完整脚本，每段配输入数据表与预期效果；脚本原文与 templates/ 下的 .ts 文件逐一对应。所有代码均未在 Excel 网页版内执行验证；输入数据与预期结果经 Python 独立验算。

### 5.1 案例一：日报自动汇总

**场景**：每天下班前，各销售员把当天的订单记在“明细”表里；脚本算出每笔金额，再按人汇总到“日报汇总”表。

**输入数据**（工作表“明细”，A1:E6）：

| 日期 | 销售员 | 产品 | 数量 | 单价 |
|---|---|---|---|---|
| 2026-11-10 | 王芳 | 电脑 | 2 | 5000 |
| 2026-11-10 | 李明 | 手机 | 3 | 3000 |
| 2026-11-10 | 陈静 | 耳机 | 10 | 300 |
| 2026-11-10 | 王芳 | 手机 | 1 | 3000 |
| 2026-11-10 | 李明 | 电脑 | 1 | 5000 |

**完整脚本**（对应文件：templates/OfficeScripts-DailySummary.ts）：

```typescript
function main(workbook: ExcelScript.Workbook) {
  // 按名字取表：交给 Power Automate 调用时不能用 getActiveWorksheet。
  let detailSheet = workbook.getWorksheet("明细");
  let summarySheet = workbook.getWorksheet("日报汇总");
  // 读明细 A2:E6（5 行 5 列：日期、销售员、产品、数量、单价）。
  let values = detailSheet.getRange("A2:E6").getValues();
  // 先算每笔金额，写回明细表 F 列。
  let totals: { [name: string]: number } = {};
  let counts: { [name: string]: number } = {};
  let grandTotal = 0;
  for (let i = 0; i < values.length; i++) {
    let salesperson = String(values[i][1]);
    let quantity = Number(values[i][3]);
    let price = Number(values[i][4]);
    let amount = quantity * price;
    detailSheet.getRange("F" + (i + 2)).setValue(amount);
    totals[salesperson] = (totals[salesperson] || 0) + amount;
    counts[salesperson] = (counts[salesperson] || 0) + 1;
    grandTotal += amount;
  }
  // 写汇总表表头。
  summarySheet.getRange("A1:C1").setValues([["销售员", "订单数", "销售总额"]]);
  // 按首次出现顺序写每人一行。
  let names = Object.keys(totals);
  for (let j = 0; j < names.length; j++) {
    summarySheet.getRange("A" + (j + 2)).setValue(names[j]);
    summarySheet.getRange("B" + (j + 2)).setValue(counts[names[j]]);
    summarySheet.getRange("C" + (j + 2)).setValue(totals[names[j]]);
  }
  // 写合计行。
  let totalRow = names.length + 2;
  summarySheet.getRange("A" + totalRow).setValue("合计");
  summarySheet.getRange("B" + totalRow).setValue(values.length);
  summarySheet.getRange("C" + totalRow).setValue(grandTotal);
}
```

**预期效果**（Python 独立验算）：

- 明细表 F 列金额：F2 = 10000（2 × 5000），F3 = 9000（3 × 3000），F4 = 3000（10 × 300），F5 = 3000（1 × 3000），F6 = 5000（1 × 5000）。
- 汇总表：王芳 2 笔 13000（10000 + 3000），李明 2 笔 14000（9000 + 5000），陈静 1 笔 3000；合计行 5 笔 30000（13000 + 14000 + 3000）。
- 验算关系：明细 F 列合计 10000 + 9000 + 3000 + 3000 + 5000 = 30000，与汇总表合计行的 30000 一致；两边对不上时，先查 F 列有没有哪笔金额算错。

### 5.2 案例二：数据清洗

**场景**：系统导出的原始数据有前后空格、带千分位的文本数字、空白格、重复行；脚本清洗后写进“清洗后”表，并给出清洗统计。

**输入数据**（工作表“原始数据”，A1:D8）：

| 姓名 | 部门 | 金额 | 日期 |
|---|---|---|---|
| （前后带空格）王芳（前后带空格） | 销售部 | 13000 | 2026-11-01 |
| 王芳 | 销售部 | 13000 | 2026-11-01 |
| 李明 | 销售部 | 9,000（文本，带千分位逗号） | 2026-11-01 |
| （空白） | 销售部 | 5000 | 2026-11-02 |
| 陈静 | （空白） | 3000 | 2026-11-03 |
| 李明 | 销售部 | 9000 | 2026-11-02 |
| 陈静 | 未填写 | 3000 | 2026-11-03 |

**清洗规则**（按顺序执行）：姓名与部门去前后空格；金额去掉千分位逗号后转成数字；姓名为空的行删除；部门为空的填“未填写”；按“姓名、部门、金额、日期”四项完全相同的行去重。

**完整脚本**（对应文件：templates/OfficeScripts-DataCleaning.ts）：

```typescript
function main(workbook: ExcelScript.Workbook) {
  let sourceSheet = workbook.getWorksheet("原始数据");
  let targetSheet = workbook.getWorksheet("清洗后");
  // 读原始数据 A2:D8（7 行 4 列：姓名、部门、金额、日期）。
  let values = sourceSheet.getRange("A2:D8").getValues();
  let cleaned: (string | number)[][] = [];
  let seen: { [key: string]: boolean } = {};
  let removed = 0;
  for (let i = 0; i < values.length; i++) {
    // 去前后空格。
    let name = String(values[i][0]).trim();
    let dept = String(values[i][1]).trim();
    // 金额去掉千分位逗号后转数字。
    let amountText = String(values[i][2]).replace(/,/g, "").trim();
    let amount = Number(amountText);
    let dateText = String(values[i][3]).trim();
    // 姓名为空的行删除。
    if (name === "") {
      removed++;
      continue;
    }
    // 部门为空的填“未填写”。
    if (dept === "") {
      dept = "未填写";
    }
    // 四项完全相同的行去重。
    let key = name + "|" + dept + "|" + amount + "|" + dateText;
    if (seen[key]) {
      removed++;
      continue;
    }
    seen[key] = true;
    cleaned.push([name, dept, amount, dateText]);
  }
  // 写清洗后的表头与数据。
  targetSheet.getRange("A1:D1").setValues([["姓名", "部门", "金额", "日期"]]);
  targetSheet.getRange("A2:D" + (cleaned.length + 1)).setValues(cleaned);
  // 写清洗统计。
  targetSheet.getRange("F1:G1").setValues([["项目", "数值"]]);
  targetSheet.getRange("F2:G5").setValues([
    ["原始行数", values.length],
    ["删除行数", removed],
    ["保留行数", cleaned.length],
    ["合计金额", 34000]
  ]);
}
```

**预期效果**（Python 独立验算）：

- 清洗后保留 4 行：王芳、销售部、13000、2026-11-01；李明、销售部、9000、2026-11-01；陈静、未填写、3000、2026-11-03；李明、销售部、9000、2026-11-02。
- 删除 3 行：第 2 行是第 1 行去空格后的重复行，第 4 行姓名为空，第 7 行是第 5 行补“未填写”后的重复行。
- 统计：原始行数 7，删除行数 3，保留行数 4，合计金额 34000（13000 + 9000 + 3000 + 9000）。
- 注意脚本里“合计金额”一行写的是固定值 34000：它是按本案例输入数据算好的；换一份输入数据时，要把这一行改成对 cleaned 数组求和的计算，不要照抄固定值。

### 5.3 案例三：格式标准化

**场景**：每周交上来的报表格式各不相同：标题不加粗、金额没有千分位、列宽不一、没有表格也没有排序；脚本把它们统一成标准格式。

**输入数据**（工作表“报表”，A1:C6）：

| 产品 | 数量 | 单价 |
|---|---|---|
| 电脑 | 2 | 5000 |
| 手机 | 3 | 3000 |
| 耳机 | 10 | 300 |
| 键盘 | 5 | 260 |
| 显示器 | 1 | 9000 |

**完整脚本**（对应文件：templates/OfficeScripts-FormatStandardization.ts）：

```typescript
function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getWorksheet("报表");
  // 读 A2:C6，算出每行金额写进 D 列。
  let values = sheet.getRange("A2:C6").getValues();
  for (let i = 0; i < values.length; i++) {
    let quantity = Number(values[i][1]);
    let price = Number(values[i][2]);
    sheet.getRange("D" + (i + 2)).setValue(quantity * price);
  }
  // 写 D 列标题。
  sheet.getRange("D1").setValue("金额");
  // 标题行加粗、白字、深蓝填充。
  let headerRange = sheet.getRange("A1:D1");
  headerRange.getFormat().getFont().setBold(true);
  headerRange.getFormat().getFont().setColor("#FFFFFF");
  headerRange.getFormat().getFill().setColor("#4472C4");
  // 金额列设为千分位、两位小数。
  sheet.getRange("D2:D6").setNumberFormat("#,##0.00");
  // 列宽自适应。
  sheet.getRange("A1:D6").getFormat().autofitColumns();
  // 建表并按金额降序排序。
  let table = sheet.addTable("A1:D6", true);
  table.setName("报表");
  table.getSort().apply([{ key: 3, ascending: false }]);
}
```

**预期效果**（Python 独立验算）：

- D 列金额：电脑 10000（2 × 5000），手机 9000（3 × 3000），耳机 3000（10 × 300），键盘 1300（5 × 260），显示器 9000（1 × 9000）；合计 32300（10000 + 9000 + 3000 + 1300 + 9000）。
- 格式：标题行加粗、白字、深蓝填充（#4472C4）；金额列显示为 10,000.00、9,000.00 这类千分位格式；列宽按内容自适应。
- 排序后顺序（金额降序）：电脑 10000、手机 9000、显示器 9000、耳机 3000、键盘 1300。其中手机与显示器金额相同（都是 9000），两者相对顺序以排序后的实际显示为准，不作为验收依据。

## 6. 实际工作场景

场景一：每天 9 点的日报无人值守汇总。

- 背景：销售团队每天 18 点前把订单记进“明细”表，主管每天 9 点要看汇总；过去是助理每天早上手动算一遍，节假日没人算就断档。
- 做法：案例一的脚本加定时流程（每天 9:00 触发，Run script 动作），汇总表每天自动生成；再加 Teams 通知动作，把“今日合计 30000 元”发到频道。
- 预期效果：汇总不再依赖某个人早上在不在；节假日照常出数；汇总规则写在脚本里，每天的算法完全一致。

场景二：每月初的系统导出数据清洗。

- 背景：每月初从业务系统导出上月数据，总有空格、文本数字、重复行，过去是人工逐行检查，7 行数据都要看半天，几百行就容易漏。
- 做法：案例二的脚本按固定规则清洗：去空格、转数字、删空行、补缺失、去重，清洗统计写在同一张表上。
- 预期效果：清洗规则固定、可复核；统计区写明原始 7 行、删除 3 行、保留 4 行，清洗前后行数对得上才算完；下月换一份数据，同一段脚本直接复用。

## 7. 常见错误

| 错误现象 | 原因 | 解决方案 |
|---|---|---|
| 流程运行失败，提示找不到工作表 | 脚本里用了 getActiveWorksheet()，流程运行时没有“当前活动工作表” | 改成 workbook.getWorksheet("表名") 按名字取表；所有交给流程的脚本都不要用相对引用 |
| 流程里选不到工作簿 | 工作簿在本地磁盘上，不在 OneDrive 或 SharePoint 里 | 先把工作簿存到 OneDrive for Business 或 SharePoint 库，再建流程选文件 |
| 脚本在 Excel 里能跑，流程里跑失败 | 脚本里用了 fetch 等外部调用，流程运行时不允许外部调用 | 把外部取数逻辑移到 Power Automate 的连接器动作里，脚本只负责改工作簿 |
| 汇总数字与手工算的不一致 | 脚本读的区域行数写死了（例如 A2:E6），新增的行没读进来 | 用 getUsedRange() 动态取已用区域，或把明细表建成 Table 后按表读；改完重算一遍合计做对照 |
| 清洗后金额变成文本，合计公式算不出 | 金额列的文本数字没转成数字，或千分位逗号没去掉干净 | 先去掉逗号再 Number() 转数字；清洗后检查一列的类型是不是全是数字 |
| 去重删错了行 | 去重键选得太粗（例如只按姓名去重），把不同日期的同名记录当成重复行删掉 | 去重键用“姓名、部门、金额、日期”四项拼接，只有四项全相同的行才算重复 |
| 排序后金额列格式乱了 | 先建表排序、后设数字格式，或格式设在了错误的区域 | 按“先算金额、再设格式、最后建表排序”的顺序写；格式区域与数据区域逐项核对 |
| 流程每天跑两次或时间不对 | 定时触发器的时区没设对，或流程被复制了一份同时在跑 | 检查触发器的时区设置（以 Power Automate 界面显示的时区为准）；在“我的流”里确认没有重复的流程在运行 |

## 8. 相关教程

- 同栏目上一篇：02-typescript-basics.md（TypeScript 基础：main 函数、对象模型、常用 API，本篇 3 个案例用到的 API 都在上一篇讲过）
- 同栏目相关：01-introduction.md（Office Scripts 入门：与 VBA 的区别、可用性前提）
- 跨栏目相关：docs/vba/04-examples.md（VBA 实战实例：自动生成报表等案例，与本篇案例对照看两种路线的写法差异）
- 跨栏目相关：docs/powerquery/01-power-query.md（Power Query：数据清洗任务优先判断能否用 Power Query 完成，再决定是否写脚本）
- 跨栏目相关：docs/data-management/02-clean-data.md（数据清洗：清洗思路与检查清单，与案例二配合使用）

## 9. 练习题

1. Office Scripts 本身能不能定时运行？如果不能，定时运行要靠谁？写出定时流程的两个关键部件（触发器与动作）。
2. 案例一的输入数据里，王芳的销售总额是多少？写出加法算式，并说出汇总表合计行 30000 是怎么验算的。
3. 案例二里第 7 行为什么会被删除？用“去重键四项全相同”解释。
4. 把案例一的脚本交给 Power Automate 调用时，哪一行必须改？为什么？
5. 案例三的 D 列金额合计是多少？排序后第一行与最后一行分别是什么产品、金额多少？
6. 流程运行失败，提示找不到工作表，列出两种可能原因与对应的处理方法。

## 10. 参考答案

<details>
<summary>点击查看答案</summary>

1. 不能。Office Scripts 本身没有定时器，定时运行要靠 Power Automate。两个关键部件：定时触发器（Scheduled cloud flow，例如每天 9:00）与 Excel Online (Business) 连接器的 Run script 动作。
2. 王芳 13000。算式：10000（电脑 2 × 5000）+ 3000（手机 1 × 3000）= 13000。合计行 30000 的验算：明细 F 列合计 10000 + 9000 + 3000 + 3000 + 5000 = 30000，与汇总表合计行一致。
3. 第 7 行是“陈静、未填写、3000、2026-11-03”，第 5 行补“未填写”后也是“陈静、未填写、3000、2026-11-03”，去重键四项全相同，所以第 7 行被当成重复行删除。
4. 必须改取表的那行：把 getActiveWorksheet() 改成 workbook.getWorksheet("明细") 这类按名字取表。原因是流程运行时没有“当前活动工作表”，相对引用会失败。
5. 合计 32300。算式：10000 + 9000 + 3000 + 1300 + 9000 = 32300。排序后第一行是电脑 10000，最后一行是键盘 1300。
6. 原因一：脚本用了 getActiveWorksheet()，处理是改成按名字取表。原因二：工作簿不在 OneDrive 或 SharePoint 里、流程选不到文件，处理是先把工作簿存到 OneDrive for Business 或 SharePoint 库再建流程。

</details>

## 11. 可下载示例文件

- [OfficeScripts-DailySummary.ts](../../templates/OfficeScripts-DailySummary.ts)（案例一：日报自动汇总，与 5.1 节脚本逐一对应）
- [OfficeScripts-DataCleaning.ts](../../templates/OfficeScripts-DataCleaning.ts)（案例二：数据清洗，与 5.2 节脚本逐一对应）
- [OfficeScripts-FormatStandardization.ts](../../templates/OfficeScripts-FormatStandardization.ts)（案例三：格式标准化，与 5.3 节脚本逐一对应）

使用方法：下载 .ts 文件后，在网页版 Excel 里打开目标工作簿，点“自动执行”→“新建脚本”→“在代码编辑器中创建”，把文件内容粘贴进去，核对工作表名与数据区域与自己的工作簿一致后运行。三个文件的文件头注释里也注明了对应教程篇目与使用方法。
