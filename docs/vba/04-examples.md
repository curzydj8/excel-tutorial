# VBA 实战实例

本章提供 8 个完整可运行的实例，每个都可以直接复制到模块中按 `F5` 执行。建议先理解再运行，运行前**备份数据**。

---

## 实例 1：批量重命名工作表

把工作簿中所有工作表按"部门_序号"规则重命名。

```vb
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
```

**要点**：`Format(i, "00")` 把数字格式化为两位数；重名会报错，用错误处理跳过。

---

## 实例 2：合并多个工作簿

把指定文件夹下所有 Excel 文件的第一个工作表数据合并到当前工作簿（假设表头都在第 1 行，数据从第 2 行开始）。

```vb
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
```

**要点**：`Dir()` 函数遍历文件夹；`FileDialog` 让用户选文件夹；记得 `ReadOnly:=True` 打开和用完关闭。

---

## 实例 3：自动生成报表（带格式）

从"数据源"表读取销售数据，生成带标题、格式、汇总行的"报表"表。

```vb
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
```

**要点**：先删旧表再建新表；金额列用公式而不是写死值，源数据变了报表自动更新。

---

## 实例 4：查找替换增强版

普通查找替换一次只能处理一个词。这个版本批量处理"对照表"中的多组替换，并报告替换次数。

```vb
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
```

**要点**：`Replace` 的 `LookAt:=xlPart` 表示部分匹配；`CountIf` 配通配符 `*` 先统计。

---

## 实例 5：数据去重

删除指定列的重复行，保留第一次出现的，并报告删除了多少行。

```vb
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
```

**要点**：`RemoveDuplicates` 是 Excel 2007+ 的内置方法，比自己写循环快得多；`Header:=xlYes` 表示第一行是表头不参与去重。

---

## 实例 6：批量插入图片

把文件夹中的图片按文件名顺序插入到工作表，每行一张，自动调整行高列宽。

```vb
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
```

**要点**：`Pictures.Insert` 插入图片；设置 `.Placement = xlMoveAndSize` 让图片跟随单元格；先定高度再按比例算宽度避免变形。

---

## 实例 7：发送邮件（Outlook）

把报表作为附件，用 Outlook 自动发送给收件人列表。

```vb
Sub SendMailViaOutlook()
    Dim olApp As Object
    Dim olMail As Object
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim toAddr As String, subject As String, body As String
    Dim attachPath As String
    Dim sentCount As Long

    ' 收件人清单：A列=邮箱，B列=姓名
    Set ws = ThisWorkbook.Worksheets("收件人")
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    attachPath = ThisWorkbook.Path & "\报表.pdf"   ' 附件路径（需提前存在）

    ' 后期绑定：不需要引用 Outlook 库，换电脑也能跑
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    If olApp Is Nothing Then
        MsgBox "未检测到 Outlook，请先安装并配置"
        Exit Sub
    End If

    sentCount = 0
    For i = 2 To lastRow
        toAddr = Trim(CStr(ws.Cells(i, 1).Value))
        If toAddr Like "*@*" Then   ' 简单校验邮箱格式
            Set olMail = olApp.CreateItem(0)   ' 0 = 邮件
            With olMail
                .To = toAddr
                .Subject = "2026年销售报表"
                .Body = ws.Cells(i, 2).Value & "，你好：" & vbCrLf & vbCrLf _
                      & "附件是本月销售报表，请查收。" & vbCrLf & vbCrLf _
                      & "报表系统（自动发送）"
                If Dir(attachPath) <> "" Then
                    .Attachments.Add attachPath
                End If
                .Send   ' 改成 .Display 可先预览不直接发
            End With
            sentCount = sentCount + 1
        End If
    Next i

    Set olMail = Nothing
    Set olApp = Nothing
    MsgBox "发送完成，共发送 " & sentCount & " 封"
End Sub
```

**要点**：
- 用 `CreateObject("Outlook.Application")` 后期绑定，不用手动加引用。
- 测试时先用 `.Display` 预览，确认无误再改 `.Send`。
- `vbCrLf` 是换行符。
- Outlook 可能会弹安全警告，这是正常保护机制。

---

## 实例 8：自定义函数（UDF）

写在标准模块中的 `Function` 可直接在单元格公式中使用。下面是三个实用的自定义函数。

```vb
Option Explicit

' 1. 提取身份证中的出生日期
' 用法：=GetBirth(A1)
Function GetBirth(idCard As String) As String
    Dim s As String
    s = Trim(idCard)
    If Len(s) = 18 Then
        GetBirth = Mid(s, 7, 4) & "-" & Mid(s, 11, 2) & "-" & Mid(s, 13, 2)
    ElseIf Len(s) = 15 Then
        GetBirth = "19" & Mid(s, 7, 2) & "-" & Mid(s, 9, 2) & "-" & Mid(s, 11, 2)
    Else
        GetBirth = "身份证号无效"
    End If
End Function

' 2. 判断是否为周末
' 用法：=IsWeekend(A1)
Function IsWeekend(d As Date) As Boolean
    Dim w As Long
    w = Weekday(d, vbMonday)   ' 以周一为一周开始
    IsWeekend = (w >= 6)
End Function

' 3. 人民币大写转换
' 用法：=RMBUpper(A1)
Function RMBUpper(amount As Double) As String
    Dim digits As Variant, units As Variant
    Dim s As String, i As Long, result As String
    Dim n As Long, digit As Long

    digits = Array("", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖")
    units = Array("", "拾", "佰", "仟")

    If amount < 0 Or amount >= 100000 Then
        RMBUpper = "超出范围"
        Exit Function
    End If

    n = Int(amount)
    s = CStr(n)
    result = ""
    For i = 1 To Len(s)
        digit = Val(Mid(s, i, 1))
        If digit <> 0 Then
            result = result & digits(digit) & units(Len(s) - i)
        Else
            ' 避免连续多个零
            If Right(result, 1) <> "零" And i < Len(s) Then
                result = result & "零"
            End If
        End If
    Next i

    ' 处理小数部分
    Dim cents As Long
    cents = Round((amount - n) * 100)
    If cents = 0 Then
        result = result & "元整"
    Else
        Dim jiao As Long, fen As Long
        jiao = cents \ 10
        fen = cents Mod 10
        result = result & "元"
        If jiao > 0 Then result = result & digits(jiao) & "角"
        If fen > 0 Then result = result & digits(fen) & "分"
    End If

    ' 清理多余的零
    Do While InStr(result, "零零") > 0
        result = Replace(result, "零零", "零")
    Loop
    result = Replace(result, "零元", "元")
    If Left(result, 1) = "零" Then result = Mid(result, 2)

    RMBUpper = result
End Function
```

**UDF 注意事项**：
1. 必须写在**标准模块**（不是工作表模块/Sheet 代码区）。
2. 不能修改其他单元格的值/格式，只能返回一个值（可以返回数组形成动态数组）。
3. 默认不随工作表重算，可按 `Ctrl+Alt+F9` 强制重算，或在函数内加 `Application.Volatile`。
4. 名字不要与内置函数重名。

```vb
' 让函数每次重算都执行
Function MyUDF() As String
    Application.Volatile
    MyUDF = "..."
End Function
```

---

## 运行前的检查清单

1. ☐ 数据已备份（另存一份）
2. ☐ 代码在 `.xlsm` 工作簿的标准模块中
3. ☐ `Option Explicit` 已开启，编译无报错（菜单「调试 → 编译」）
4. ☐ 涉及删除/覆盖的代码，先在小范围测试
5. ☐ 邮件类代码先用 `.Display` 预览，确认后再 `.Send`
