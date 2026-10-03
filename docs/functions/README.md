# Excel 函数大全 · 索引

本目录按功能分类收录 Excel 常用函数，每个函数包含**语法、参数说明、实用示例、常见错误提醒**，示例均为原创编写。

> 版本说明：标注「365」的函数需要 Microsoft 365 / Excel 2021 及以上版本；标注年份的为引入版本；未标注的为全版本通用。

## 分类导航

| 文件 | 分类 | 函数数量 | 覆盖场景 |
|------|------|----------|----------|
| [01-math.md](01-math.md) | 数学函数 | 19 | 求和、平均、取整、随机数、条件求和 |
| [02-text.md](02-text.md) | 文本函数 | 18 | 提取、合并、替换、查找、拆分字符串 |
| [03-datetime.md](03-datetime.md) | 日期时间函数 | 16 | 年龄、工龄、工作日、到期提醒 |
| [04-logical.md](04-logical.md) | 逻辑函数 | 10 | 条件判断、错误处理、逻辑组合 |
| [05-lookup.md](05-lookup.md) | 查找引用函数 | 14 | 查表、匹配、动态数组、去重筛选 |
| [06-statistical.md](06-statistical.md) | 统计函数 | 16 | 计数、最值、排名、中位数、标准差 |
| [07-financial.md](07-financial.md) | 财务函数 | 9 | 房贷月供、定投终值、NPV、IRR、折旧 |
| **合计** | | **102** | |

## 函数速查表

### 数学函数（[01-math.md](01-math.md)）

| 函数 | 一句话说明 |
|------|-----------|
| SUM | 求和 |
| SUMIF / SUMIFS | 单条件 / 多条件求和 |
| AVERAGE / AVERAGEIF | 求平均 / 条件平均 |
| COUNT | 统计数字个数 |
| ROUND / ROUNDUP / ROUNDDOWN | 四舍五入 / 向上进位 / 向下舍去 |
| INT | 向下取整 |
| MOD | 取余数 |
| ABS | 绝对值 |
| SQRT / POWER | 平方根 / 乘方 |
| RAND / RANDBETWEEN | 随机小数 / 随机整数 |
| SUMPRODUCT | 加权求和、多条件计数 |
| SUBTOTAL | 筛选后汇总（忽略隐藏行） |
| AGGREGATE | 忽略错误值的多功能汇总 |

### 文本函数（[02-text.md](02-text.md)）

| 函数 | 一句话说明 |
|------|-----------|
| LEFT / RIGHT / MID | 左取 / 右取 / 中间取字符 |
| LEN | 字符个数 |
| CONCAT / TEXTJOIN | 合并文本 / 带分隔符合并 |
| TEXT / VALUE | 数值转文本 / 文本转数值 |
| UPPER / LOWER | 转大写 / 转小写 |
| TRIM | 去多余空格 |
| SUBSTITUTE / REPLACE | 按内容替换 / 按位置替换 |
| FIND / SEARCH | 查找位置（区分大小写 / 不区分） |
| TEXTBEFORE / TEXTAFTER / TEXTSPLIT | 取分隔符前后 / 拆分（365） |

### 日期时间函数（[03-datetime.md](03-datetime.md)）

| 函数 | 一句话说明 |
|------|-----------|
| TODAY / NOW | 今天日期 / 当前日期时间 |
| DATE / TIME | 组合日期 / 组合时间 |
| YEAR / MONTH / DAY | 取年 / 月 / 日 |
| HOUR / MINUTE | 取小时 / 分钟 |
| DATEDIF | 算年龄、工龄 |
| EDATE / EOMONTH | n 月后的日期 / 月末日期 |
| WEEKDAY | 星期几 |
| NETWORKDAYS / WORKDAY | 工作日天数 / 推算工作日 |
| DATEVALUE | 文本转日期 |

### 逻辑函数（[04-logical.md](04-logical.md)）

| 函数 | 一句话说明 |
|------|-----------|
| IF / IFS | 单条件 / 多条件判断 |
| AND / OR / NOT / XOR | 且 / 或 / 非 / 异或 |
| IFERROR / IFNA | 错误捕获 / 仅捕获 #N/A |
| LET | 命名中间变量（365） |
| SWITCH | 精确匹配多分支 |

### 查找引用函数（[05-lookup.md](05-lookup.md)）

| 函数 | 一句话说明 |
|------|-----------|
| VLOOKUP / HLOOKUP | 纵向 / 横向查找 |
| XLOOKUP | 新一代查找（365） |
| INDEX + MATCH | 灵活查表组合 |
| OFFSET / INDIRECT | 偏移引用 / 文本转引用 |
| CHOOSE | 按序号取值 |
| ROW / COLUMN | 行号 / 列号 |
| TRANSPOSE / SORT / FILTER / UNIQUE | 转置 / 排序 / 筛选 / 去重（365） |

### 统计函数（[06-statistical.md](06-statistical.md)）

| 函数 | 一句话说明 |
|------|-----------|
| COUNT / COUNTA / COUNTBLANK | 数字计数 / 非空计数 / 空计数 |
| COUNTIF / COUNTIFS | 单条件 / 多条件计数 |
| AVERAGEIF / AVERAGEIFS | 单条件 / 多条件平均 |
| MAX / MIN | 最大值 / 最小值 |
| LARGE / SMALL | 第 k 大 / 第 k 小 |
| RANK | 排名 |
| MEDIAN / MODE | 中位数 / 众数 |
| STDEV / VAR | 标准差 / 方差 |

### 财务函数（[07-financial.md](07-financial.md)）

| 函数 | 一句话说明 |
|------|-----------|
| PMT | 每期还款额（房贷月供） |
| FV / PV | 终值 / 现值 |
| NPV / IRR | 净现值 / 内部收益率 |
| RATE / NPER | 反推利率 / 反推期数 |
| SLN / DB | 直线折旧 / 余额递减折旧 |

## 学习建议

1. **新手**：先掌握 SUM、IF、VLOOKUP、COUNTIF、SUMIF 这 5 个，能解决 80% 日常需求
2. **进阶**：学习 INDEX+MATCH、SUMPRODUCT、数据透视表配合，摆脱对单一函数的依赖
3. **365 用户**：优先学 XLOOKUP、FILTER、UNIQUE、LET，很多旧套路可以扔掉了
4. 每个函数的「常见错误」都是实战中高频踩坑点，建议通读一遍再动手
