// Office Scripts 示例：格式标准化
// 对应教程：docs/office-scripts/03-automation-cases.md 第 5.3 节（案例三）
// 使用方法：在网页版 Excel 里打开工作簿，确认有“报表”工作表，
// A1:C6 按教程数据准备好后，点“自动执行”→“新建脚本”→“在代码编辑器中创建”，
// 把本文件内容粘贴进去，核对工作表名与数据区域无误后运行。
// 本脚本未在 Excel 网页版内执行验证；预期效果见教程 5.3 节（金额合计 32300，降序后首行电脑 10000，末行键盘 1300）。

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
