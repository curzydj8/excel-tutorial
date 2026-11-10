' VBA-MergeWorkbooks.bas
' 对应教程：docs/vba/04-examples.md 实例 2：合并多个工作簿
' 使用方法：在 VBE 中导入本文件，或新建标准模块后全文粘贴；运行 MergeWorkbooks，按提示选择要合并的文件夹（运行前备份）。
' 说明：本文件代码经语法扫描核对，未在 Excel 内执行验证（无 Excel 环境）；导入后请先在副本工作簿上试跑。
Attribute VB_Name = "modMergeWorkbooks"
Option Explicit

Sub MergeWorkbooks()
    Dim folder As String
    Dim fileName As String
    Dim wb As Workbook
    Dim targetRow As Long
    Dim srcRows As Long

    ' 让用户选择文件夹
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "选择要合并的文件夹"
        If .Show <> -1 Then Exit Sub
        folder = .SelectedItems(1) & "\"
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    targetRow = 2   ' 目标表第1行是表头，从第2行开始写
    fileName = Dir(folder & "*.xlsx")

    Do While fileName <> ""
        ' 跳过当前工作簿自己
        If folder & fileName <> ThisWorkbook.FullName Then
            Set wb = Workbooks.Open(folder & fileName, ReadOnly:=True)
            ' 找到源表最后一行
            srcRows = wb.Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row
            If srcRows >= 2 Then
                wb.Worksheets(1).Range("A2:Z" & srcRows).Copy _
                    Destination:=ThisWorkbook.Worksheets(1).Cells(targetRow, 1)
                targetRow = targetRow + (srcRows - 1)
            End If
            wb.Close SaveChanges:=False
        End If
        fileName = Dir()   ' 下一个文件
    Loop

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合并完成，共写入 " & (targetRow - 2) & " 行数据"
End Sub
