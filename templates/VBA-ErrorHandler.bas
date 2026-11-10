' VBA-ErrorHandler.bas
' 对应教程：docs/vba/05-advanced.md 第 4.4 节：错误处理进阶（集中处理与日志）
' 使用方法：在 VBE 中导入本文件，或新建标准模块后全文粘贴；业务过程按文末 OpenReport 的固定写法（顶部 On Error GoTo ErrHandler、文末 ErrHandler 段调 HandleError）接入；错误日志写到工作簿同目录的 ErrorLog.txt。
' 说明：本文件代码经语法扫描核对，未在 Excel 内执行验证（无 Excel 环境）；导入后请先在副本工作簿上试跑。
Attribute VB_Name = "modErrorHandler"
' ---- 以下代码放在标准模块 modErrorHandler ----
Option Explicit

' 写日志：追加到工作簿同目录的 ErrorLog.txt（每行一条）
Public Sub LogError(ByVal procName As String)
    Dim f As Long
    Dim logPath As String

    logPath = ThisWorkbook.Path & "\ErrorLog.txt"
    f = FreeFile
    Open logPath For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & procName _
        & " | " & Err.Number & " | " & Err.Description
    Close #f
End Sub

' 集中处理：按错误号分支提示，再记日志
Public Sub HandleError(ByVal procName As String)
    Select Case Err.Number
        Case 9
            MsgBox "下标越界：要找的工作表或数组元素不存在，请检查名称与范围。", vbExclamation
        Case 13
            MsgBox "类型不匹配：请检查输入的数据类型（如把文字填进了数字格）。", vbExclamation
        Case 1004
            MsgBox "文件不存在或路径错误，请检查路径后重试。", vbExclamation
        Case Else
            MsgBox "出错了（" & Err.Number & "）：" & Err.Description & "，已记入错误日志。", vbCritical
    End Select
    LogError procName
    Err.Clear
End Sub

' ---- 业务过程的固定写法：顶部一行，文末一段 ----
Option Explicit

Sub OpenReport()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Workbooks.Open "D:\不存在.xlsx"   ' 演示用：该文件不存在，触发 1004

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True   ' 恢复现场：出错也要把开关开回来
    HandleError "OpenReport"
End Sub
