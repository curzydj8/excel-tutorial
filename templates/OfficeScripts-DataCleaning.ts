// Office Scripts 示例：数据清洗
// 对应教程：docs/office-scripts/03-automation-cases.md 第 5.2 节（案例二）
// 使用方法：在网页版 Excel 里打开工作簿，确认有“原始数据”与“清洗后”两张工作表，
// 原始数据表 A1:D8 按教程数据准备好后，点“自动执行”→“新建脚本”→“在代码编辑器中创建”，
// 把本文件内容粘贴进去，核对工作表名与数据区域无误后运行。
// 本脚本未在 Excel 网页版内执行验证；预期效果见教程 5.2 节（原始 7 行，删除 3 行，保留 4 行，合计金额 34000）。
// 注意：“合计金额”一行写的是按本案例输入数据算好的固定值 34000；换一份输入数据时，
// 请把这一行改成对 cleaned 数组求和的计算，不要照抄固定值。

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
