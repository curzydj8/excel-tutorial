' VBA-BatchReplace.bas
' 对应教程：docs/vba/04-examples.md 实例 4：查找替换增强版
' 使用方法：在 VBE 中导入本文件，或新建标准模块后全文粘贴；先建好“对照表”（A列查找内容、B列替换为）与“数据”表，再运行 BatchReplace。
' 说明：本文件代码经语法扫描核对，未在 Excel 内执行验证（无 Excel 环境）；导入后请先在副本工作簿上试跑。
Attribute VB_Name = "modBatchReplace"
Option Explicit

Sub BatchReplace()
    Dim mapWs As Worksheet, dataWs As Worksheet
    Dim lastMap As Long, i As Long
    Dim findStr As String, replStr As String
    Dim count As Long, totalCount As Long

    ' 对照表：A列=查找内容，B列=替换为
    Set mapWs = ThisWorkbook.Worksheets("对照表")
    Set dataWs = ThisWorkbook.Worksheets("数据")

    lastMap = mapWs.Cells(Rows.Count, 1).End(xlUp).Row
    totalCount = 0

    Application.ScreenUpdating = False
    For i = 2 To lastMap
        findStr = CStr(mapWs.Cells(i, 1).Value)
        replStr = CStr(mapWs.Cells(i, 2).Value)
        If findStr <> "" Then
            ' 先数有多少个，再替换
            count = Application.WorksheetFunction.CountIf(dataWs.UsedRange, "*" & findStr & "*")
            dataWs.UsedRange.Replace What:=findStr, Replacement:=replStr, _
                LookAt:=xlPart, MatchCase:=False
            totalCount = totalCount + count
            Debug.Print findStr & " -> " & replStr & "（约" & count & "处）"
        End If
    Next i
    Application.ScreenUpdating = True

    MsgBox "批量替换完成，共处理 " & (lastMap - 1) & " 组规则"
End Sub
