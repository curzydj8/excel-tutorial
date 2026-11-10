' VBA-BatchRename.bas
' 对应教程：docs/vba/04-examples.md 实例 1：批量重命名工作表
' 使用方法：在 VBE 中「文件 → 导入文件」导入本文件，或新建标准模块后全文粘贴；运行 BatchRenameSheets（运行前备份工作簿）。
' 说明：本文件代码经语法扫描核对，未在 Excel 内执行验证（无 Excel 环境）；导入后请先在副本工作簿上试跑。
Attribute VB_Name = "modBatchRename"
Option Explicit

Sub BatchRenameSheets()
    Dim ws As Worksheet
    Dim i As Long
    Dim newName As String

    Application.DisplayAlerts = False   ' 重命名冲突时不弹窗，由代码处理
    i = 1
    For Each ws In ThisWorkbook.Worksheets
        newName = "部门_" & Format(i, "00")   ' 部门_01、部门_02...
        On Error Resume Next                  ' 重名时跳过
        ws.Name = newName
        If Err.Number <> 0 Then
            Debug.Print "跳过重名：" & ws.Name
            Err.Clear
        End If
        On Error GoTo 0
        i = i + 1
    Next ws
    Application.DisplayAlerts = True

    MsgBox "重命名完成，共处理 " & (i - 1) & " 个工作表"
End Sub
