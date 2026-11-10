' VBA-UserForm-Demo.bas
' 对应教程：docs/vba/05-advanced.md 第 4.1 节：用户窗体（UserForm）
' 使用方法：本文件分两部分：A 部分（窗体代码）双击窗体 frmEntry 打开代码窗口后粘贴；B 部分（ShowEntryForm）放在标准模块，可直接导入或粘贴后按 F5 运行。窗体本身需在 VBE 中「插入 → 用户窗体」新建并按教程第 2.1 节放好控件。
' 说明：本文件代码经语法扫描核对，未在 Excel 内执行验证（无 Excel 环境）；导入后请先在副本工作簿上试跑。
' ===== A 部分：窗体 frmEntry 的代码（粘贴进窗体代码窗口）=====
' ---- 以下代码粘贴到窗体 frmEntry 的代码窗口（双击窗体打开）----
Option Explicit

Private Sub UserForm_Initialize()
    ' 窗体第一次加载时执行一次：填下拉框初始值
    With Me.cmbType
        .AddItem "差旅"
        .AddItem "办公"
        .AddItem "招待"
        .ListIndex = 0   ' 默认选中第一项
    End With
    ' 注意：不要在这里写 txtName.SetFocus（窗体还没显示，设焦点会报错）；
    ' 想让姓名框默认获得焦点，把 txtName 的 TabIndex 属性设为 0 即可
End Sub

Private Sub cmdOK_Click()
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim amount As Double

    ' --- 校验 ---
    If Trim(Me.txtName.Text) = "" Then
        MsgBox "姓名不能为空", vbExclamation
        Me.txtName.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(Me.txtAmount.Text) Then
        MsgBox "金额必须是数字", vbExclamation
        Me.txtAmount.SetFocus
        Exit Sub
    End If
    amount = CDbl(Me.txtAmount.Text)
    If amount <= 0 Then
        MsgBox "金额必须大于 0", vbExclamation
        Me.txtAmount.SetFocus
        Exit Sub
    End If

    ' --- 写入 ---
    Set ws = ThisWorkbook.Worksheets("报销")
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Trim(Me.txtName.Text)
    ws.Cells(nextRow, 2).Value = amount
    ws.Cells(nextRow, 3).Value = Me.cmbType.Value
    ws.Cells(nextRow, 4).Value = Date

    Unload Me   ' 彻底卸载，输入值清空
End Sub

Private Sub cmdCancel_Click()
    Unload Me   ' 不写数据，直接关闭
End Sub

' ===== B 部分：标准模块代码（显示窗体）=====
' ---- 以下代码放在标准模块（如 modMain），用来显示窗体 ----
Option Explicit

Sub ShowEntryForm()
    frmEntry.Show   ' 默认模态（vbModal）：关闭窗体前不能操作工作表
    ' frmEntry.Show vbModeless   ' 改成这一行则为非模态：可一边开窗体一边操作表格
End Sub
