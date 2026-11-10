' VBA-ReportGenerator.bas
' 对应教程：docs/vba/04-examples.md 实例 3：自动生成报表（带格式）
' 使用方法：在 VBE 中导入本文件，或新建标准模块后全文粘贴；先建好名为“数据源”的工作表（A列产品、B列数量、C列单价，第 1 行为表头），再运行 GenerateReport。
' 说明：本文件代码经语法扫描核对，未在 Excel 内执行验证（无 Excel 环境）；导入后请先在副本工作簿上试跑。
Attribute VB_Name = "modReportGenerator"
Option Explicit

Sub GenerateReport()
    Dim src As Worksheet, rpt As Worksheet
    Dim lastRow As Long, i As Long
    Dim total As Double

    Set src = ThisWorkbook.Worksheets("数据源")

    ' 删除旧报表（如果存在）
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("报表").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set rpt = ThisWorkbook.Worksheets.Add(After:=src)
    rpt.Name = "报表"

    ' --- 标题区 ---
    With rpt.Range("A1:D1")
        .Merge
        .Value = "2026年销售报表（" & Format(Date, "yyyy-mm-dd") & " 生成）"
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
    End With
    rpt.Range("A1").RowHeight = 30

    ' --- 表头 ---
    Dim headers As Variant
    headers = Array("产品", "数量", "单价", "金额")
    For i = 0 To 3
        With rpt.Cells(3, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .Interior.Color = RGB(217, 217, 217)
            .HorizontalAlignment = xlCenter
        End With
    Next i

    ' --- 数据区 ---
    lastRow = src.Cells(Rows.Count, 1).End(xlUp).Row
    total = 0
    For i = 2 To lastRow
        rpt.Cells(i + 2, 1).Value = src.Cells(i, 1).Value   ' 产品
        rpt.Cells(i + 2, 2).Value = src.Cells(i, 2).Value   ' 数量
        rpt.Cells(i + 2, 3).Value = src.Cells(i, 3).Value   ' 单价
        rpt.Cells(i + 2, 4).Formula = "=B" & (i + 2) & "*C" & (i + 2)  ' 金额公式
        total = total + src.Cells(i, 2).Value * src.Cells(i, 3).Value
    Next i

    ' --- 汇总行 ---
    Dim sumRow As Long
    sumRow = lastRow + 3
    With rpt.Range("A" & sumRow & ":C" & sumRow)
        .Merge
        .Value = "合计"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    rpt.Cells(sumRow, 4).Value = total
    rpt.Cells(sumRow, 4).Font.Bold = True
    rpt.Cells(sumRow, 4).NumberFormat = "#,##0.00"

    ' --- 收尾 ---
    rpt.Range("A3:D" & sumRow).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    rpt.Range("A3:D" & sumRow).Borders.LineStyle = xlContinuous
    rpt.Columns("A:D").AutoFit

    MsgBox "报表已生成，共 " & (lastRow - 1) & " 条数据，合计 " & Format(total, "#,##0.00")
End Sub
