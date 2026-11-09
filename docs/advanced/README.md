# 🚀 高级方法

- **适用场景**：已掌握基础函数，想把"会用 Excel"变成"用 Excel 提效"的人
- **难度**：★★☆☆☆ ~ ★★★★☆（见各篇标注）

---

## 学习路径

建议按以下顺序学，每一步都是下一步的地基：

### 第一阶段：让表格"活起来"（1 周）

| 顺序 | 文档 | 学完你会 |
|------|------|----------|
| 1 | [快捷键大全](../basics/04-shortcuts.md) | 脱离鼠标，操作速度翻倍（先记 20 个标 ★ 的） |
| 2 | [条件格式](../conditional-formatting/01-conditional-formatting.md) | 异常自动标红、到期自动提醒，报表自己说话 |

### 第二阶段：报表自动化（2 周）

| 顺序 | 文档 | 学完你会 |
|------|------|----------|
| 3 | [数据验证与保护](../data-management/04-data-validation.md) | 做防呆录入模板，下拉+联动+锁公式，收表零清洗 |
| 4 | [数据透视表](../pivottable/01-pivottable.md) | 5000 行明细 3 秒出交叉汇总，环比、占比一键算 |

### 第三阶段：数据工程（2~3 周）

| 顺序 | 文档 | 学完你会 |
|------|------|----------|
| 5 | [数组公式与动态数组](../formulas/04-array-formulas.md) | 一公式返回多结果，FILTER/SORT/UNIQUE 替代手工筛选 |
| 6 | [Power Query](powerquery.md) | 月报清洗流水线：多文件合并、逆透视、自动刷新 |

学完六篇，你就具备了"**原始数据 → 自动出报表**"的完整能力链。

---

## 速查索引：按问题找文档

| 我想… | 看这里 |
|-------|--------|
| 快速选中整列数据 | [快捷键-选择](../basics/04-shortcuts.md#二选择13-个) |
| 整行标红/隔行变色 | [条件格式-公式规则](../conditional-formatting/01-conditional-formatting.md#51-l2-2-公式规则的心法) |
| 合同到期前 30 天提醒我 | [条件格式-到期提醒](../conditional-formatting/01-conditional-formatting.md#53-l2-2-新增示例逐一验算) |
| 做省市二级下拉 | [数据验证-二级联动](../data-management/04-data-validation.md#2-下拉列表最常用) |
| 防止别人改我的公式 | [数据验证-工作表保护](../data-management/04-data-validation.md#5-工作表保护锁公式开放录入区) |
| 按月统计每个销售员的业绩 | [数据透视表-实战](../pivottable/01-pivottable.md#6-实际工作场景l2-6-透视表实战6-个任务及参考答案) |
| 算环比增长率 | [数据透视表-差异百分比](../pivottable/01-pivottable.md#23-值汇总方式) |
| 提取不重复的客户名单 | [数组公式-UNIQUE](../formulas/04-array-formulas.md#unique去重) |
| 按条件自动列清单 | [数组公式-FILTER](../formulas/04-array-formulas.md#filter按条件筛选) |
| 公式报 #溢出! | [数组公式-溢出处理](../formulas/04-array-formulas.md#4-溢出-错误处理) |
| 12 个月文件合并成一张表 | [Power Query-追加查询](powerquery.md#5-追加查询纵向堆叠) |
| 交叉表转成明细表 | [Power Query-逆透视](powerquery.md#6-透视列与逆透视列) |
| 给订单匹配客户名称 | [Power Query-合并查询](powerquery.md#4-合并查询比-vlookup-更强的关联) |
| 引用透视表里的数做二次计算 | [数据透视表-GETPIVOTDATA](../pivottable/01-pivottable.md#51-getpivotdata-函数) |

---

## 各篇难度一览

| 文档 | 难度 | 预计学习时间 |
|------|------|--------------|
| 快捷键大全 | ★☆☆☆☆ | 随用随查 |
| 条件格式 | ★★☆☆☆ | 2 小时 |
| 数据验证与保护 | ★★☆☆☆ | 2 小时 |
| 数据透视表 | ★★★☆☆ | 半天 |
| 数组公式与动态数组 | ★★★★☆ | 1 天（需 365） |
| Power Query | ★★★☆☆（M 语言 ★★★★☆） | 2 天 |

## 版本提示

- 动态数组函数（FILTER/SORT/UNIQUE 等）需要 **Microsoft 365 或 Excel 2021**。
- 其余内容 Excel 2016 均可使用；切片器/时间线需 2013+。
- 公司电脑如果是旧版本，优先学透视表 + 条件格式 + 数据验证，这三样版本兼容性最好。

---

## 目录调整说明（2026-11-09，阶段 3；2026-11-10，阶段 5）

本目录已按设计 §9.1 调整：快捷键篇移至 `docs/basics/04-shortcuts.md`，数组公式篇移至 `docs/formulas/04-array-formulas.md`，数据验证篇移至 `docs/data-management/04-data-validation.md`（旧路径留重定向索引一个版本周期）；Excel Canvas 草稿已从仓库移除，继续冻结，等实机验证通过并经 Zhang 批准后再议。阶段 5：条件格式篇移至 `docs/conditional-formatting/01-conditional-formatting.md`，数据透视表篇移至 `docs/pivottable/01-pivottable.md`（旧路径留重定向索引，阶段 6 验收时删除）。本目录现仅保留 Power Query 一篇正文。

## 学习目标

学完本目录的 Power Query 篇，你将能够：用 Power Query 把月报清洗做成自动刷新的流水线。条件格式与数据透视表两篇已分别移至 `docs/conditional-formatting/` 与 `docs/pivottable/`，每篇末尾附练习题与参考答案。
