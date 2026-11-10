' VBA-BatchInsertPictures.bas
' 对应教程：docs/vba/04-examples.md 实例 6：批量插入图片
' 使用方法：在 VBE 中导入本文件，或新建标准模块后全文粘贴；运行 BatchInsertPictures，按提示选择图片文件夹。
' 说明：本文件代码经语法扫描核对，未在 Excel 内执行验证（无 Excel 环境）；导入后请先在副本工作簿上试跑。
Attribute VB_Name = "modBatchInsertPictures"
Option Explicit

Sub BatchInsertPictures()
    Dim folder As String
    Dim fileName As String
    Dim pic As Picture
    Dim row As Long
    Dim ext As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "选择图片文件夹"
        If .Show <> -1 Then Exit Sub
        folder = .SelectedItems(1) & "\"
    End With

    Application.ScreenUpdating = False
    row = 2
    ' A1写表头
    Range("A1").Value = "图片"
    Range("B1").Value = "文件名"

    fileName = Dir(folder & "*.*")
    Do While fileName <> ""
        ext = LCase(Right(fileName, 4))
        If ext = ".jpg" Or ext = "jpeg" Or ext = ".png" Or ext = ".bmp" Or ext = ".gif" Then
            ' 插入图片
            Set pic = ActiveSheet.Pictures.Insert(folder & fileName)
            With pic
                .Top = Cells(row, 1).Top + 2
                .Left = Cells(row, 1).Left + 2
                .Height = 80
                ' 按比例缩放宽度
                .Width = .Width * (80 / .Height)
                .Placement = xlMoveAndSize   ' 图片随单元格移动和缩放
            End With
            Cells(row, 1).RowHeight = 84
            Cells(row, 2).Value = fileName
            row = row + 1
        End If
        fileName = Dir()
    Loop

    Columns("A").ColumnWidth = 20
    Application.ScreenUpdating = True
    MsgBox "共插入 " & (row - 2) & " 张图片"
End Sub
