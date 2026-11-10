' VBA-RemoveDuplicates.bas
' 对应教程：docs/vba/04-examples.md 实例 5：数据去重
' 使用方法：在 VBE 中导入本文件，或新建标准模块后全文粘贴；选中要去重的工作表（按 A 列去重，第 1 行为表头），运行 RemoveDuplicates（运行前备份）。
' 说明：本文件代码经语法扫描核对，未在 Excel 内执行验证（无 Excel 环境）；导入后请先在副本工作簿上试跑。
Attribute VB_Name = "modRemoveDuplicates"
Option Explicit

Sub RemoveDuplicates()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim beforeCount As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    beforeCount = lastRow

    If lastRow < 2 Then
        MsgBox "没有数据可处理"
        Exit Sub
    End If

    ' 按A列去重（多列去重传数组，如 Array(1, 2) 表示按A、B列组合去重）
    ws.Range("A1:Z" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox "去重完成，删除了 " & (beforeCount - lastRow) & " 行重复数据"
End Sub
