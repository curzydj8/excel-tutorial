// Office Scripts 示例：日报自动汇总
// 对应教程：docs/office-scripts/03-automation-cases.md 第 5.1 节（案例一）
// 使用方法：在网页版 Excel 里打开日报工作簿，确认有“明细”与“日报汇总”两张工作表，
// 明细表 A1:E6 按教程数据准备好后，点“自动执行”→“新建脚本”→“在代码编辑器中创建”，
// 把本文件内容粘贴进去，核对工作表名与数据区域无误后运行。
// 本脚本未在 Excel 网页版内执行验证；预期效果见教程 5.1 节（汇总合计 30000）。

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
