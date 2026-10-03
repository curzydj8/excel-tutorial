# VBA 对象模型

本章讲解 Excel VBA 的核心：对象模型。理解 Application → Workbook → Worksheet → Range 这条层级链，就掌握了操作 Excel 的钥匙。

## 1. 四大对象与层级关系

Excel 的对象是树状层级结构：

```
Application（Excel 应用程序本身）
 └─ Workbooks（所有打开的工作簿集合）
     └─ Workbook（单个工作簿，如 ThisWorkbook）
         └─ Worksheets（工作表集合）
             └─ Worksheet（单个工作表，如 Sheet1）
                 └─ Range（单元格区域，如 A1:B10）
```

每一层都可以通过父对象拿到子对象：

```vb
Sub Hierarchy()
    ' 应用程序
    Debug.Print Application.Version          ' Excel 版本号

    ' 工作簿
    Debug.Print ThisWorkbook.Name            ' 当前代码所在的工作簿
    Debug.Print ActiveWorkbook.Name          ' 当前活动的工作簿
    Debug.Print Workbooks.Count              ' 打开的工作簿数量

    ' 工作表
    Debug.Print ThisWorkbook.Worksheets.Count
    Debug.Print ActiveSheet.Name             ' 当前活动的工作表

    ' 单元格
    Debug.Print ActiveSheet.Range("A1").Value
End Sub
```

> **ThisWorkbook vs ActiveWorkbook**：`ThisWorkbook` 是代码所在的工作簿（固定），`ActiveWorkbook` 是当前显示在最前面的工作簿（会变）。写代码时优先用 `ThisWorkbook`，避免用户切换窗口导致操作错对象。

## 2. Application 对象

代表 Excel 程序本身，控制全局行为：

```vb
Sub AppDemo()
    ' 提升性能三件套（大量操作前后使用）
    Application.ScreenUpdating = False   ' 关闭屏幕刷新
    Application.Calculation = xlCalculationManual  ' 关闭自动重算
    Application.EnableEvents = False     ' 关闭事件触发

    ' ... 批量操作 ...

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
```

常用成员：

| 成员 | 说明 |
|------|------|
| `ScreenUpdating` | 屏幕刷新开关，批量操作时关闭可大幅提速 |
| `Calculation` | 计算模式：`xlCalculationManual` 手动 / `xlCalculationAutomatic` 自动 |
| `EnableEvents` | 事件开关，关闭可防止 Change 事件递归触发 |
| `DisplayAlerts` | 是否显示警告框（如删除工作表确认），`False` 可静默执行 |
| `StatusBar` | 状态栏文字，可显示进度 |
| `Version` | Excel 版本 |
| `Quit` | 退出 Excel |

进度条示例：

```vb
Sub ProgressDemo()
    Dim i As Long
    Application.ScreenUpdating = False
    For i = 1 To 1000
        ' ... 处理 ...
        If i Mod 100 = 0 Then
            Application.StatusBar = "处理中 " & i & " / 1000"
        End If
    Next i
    Application.StatusBar = False   ' 恢复状态栏
    Application.ScreenUpdating = True
End Sub
```

## 3. Workbook 对象

```vb
Sub WorkbookDemo()
    Dim wb As Workbook

    ' 打开工作簿
    Set wb = Workbooks.Open("D:\数据\报表.xlsx")

    ' 新建工作簿
    Set wb = Workbooks.Add

    ' 保存 / 另存
    wb.Save
    wb.SaveAs "D:\数据\报表_备份.xlsx"

    ' 关闭（False = 不保存直接关）
    wb.Close SaveChanges:=False

    ' 遍历所有打开的工作簿
    Dim w As Workbook
    For Each w In Workbooks
        Debug.Print w.Name, w.Path
    Next w
End Sub
```

> **Set 的作用**：对象变量赋值必须用 `Set`，普通变量（数字、字符串）直接用 `=`。忘记 `Set` 会报"需要对象"错误。

## 4. Worksheet 对象

```vb
Sub WorksheetDemo()
    Dim ws As Worksheet

    ' 三种引用方式
    Set ws = ThisWorkbook.Worksheets("销售数据")  ' 按名称（推荐）
    Set ws = ThisWorkbook.Worksheets(1)           ' 按索引（第1个）
    Set ws = Sheet1                               ' 按代码名称（VBE左侧显示的名称）

    ' 新增 / 删除 / 重命名 / 复制
    ThisWorkbook.Worksheets.Add After:=ws          ' 在ws后面插入
    ws.Delete                                     ' 删除（会弹窗确认）
    ws.Name = "2026年销售"                        ' 重命名
    ws.Copy After:=ThisWorkbook.Worksheets(1)     ' 复制

    ' 隐藏 / 保护
    ws.Visible = xlSheetHidden                     ' 隐藏
    ws.Visible = xlSheetVeryHidden                  ' 深度隐藏（只能用代码取消）
    ws.Protect Password:="123"                      ' 保护工作表
    ws.Unprotect Password:="123"                    ' 取消保护
End Sub
```

删除时跳过确认框：

```vb
Application.DisplayAlerts = False
ws.Delete
Application.DisplayAlerts = True
```

## 5. Range 对象：核心中的核心

### Cells vs Range

两者都能定位单元格，区别在写法：

```vb
Sub CellsVsRange()
    ' Range：用 A1 式地址，直观
    Range("A1").Value = 1
    Range("A1:C3").Value = 2

    ' Cells：用行号+列号，适合循环
    Cells(1, 1).Value = 1        ' = Range("A1")
    Cells(1, "A").Value = 1      ' 列也可以用字母

    Dim i As Long
    For i = 1 To 10
        Cells(i, 2).Value = i * 10   ' B列填 10,20,...,100
    Next i

    ' 组合：Cells 定位起点，Resize 扩展区域
    Cells(1, 1).Resize(5, 3).Value = 0   ' A1:C5 填 0
End Sub
```

| 场景 | 推荐 |
|------|------|
| 固定地址 | `Range("A1")` |
| 循环中动态定位 | `Cells(i, j)` |
| 以某个单元格为基准扩展 | `Range("A1").Offset(2, 1)` / `.Resize(3, 2)` |

### 定位技巧

```vb
Sub RangeTips()
    ' 相对定位
    Range("B2").Offset(1, -1).Value = "x"   ' = A3（下1行，左1列）
    Range("A1").Resize(3, 2).Select         ' A1:B3

    ' 整行 / 整列
    Rows(1).Font.Bold = True                ' 第1行加粗
    Columns("B").ColumnWidth = 20            ' B列宽20

    ' 已用区域 / 当前区域
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row   ' A列最后一个非空行
    Debug.Print "数据共" & lastRow & "行"

    Range("A1").CurrentRegion.Select   ' A1所在的连续数据块

    ' UsedRange：工作表用过的区域
    ActiveSheet.UsedRange.ClearFormats
End Sub
```

`End(xlUp)` 相当于在单元格按 `Ctrl + ↑`，是找"最后一行"最常用的方法。四个方向：`xlUp`、`xlDown`、`xlToLeft`、`xlToRight`。

## 6. 常用属性

### Value / Formula

```vb
Range("A1").Value = 100              ' 设值
Range("A2").Formula = "=SUM(B1:B10)" ' 设公式（英文函数名，逗号分隔）
Range("A3").FormulaR1C1 = "=SUM(R1C2:R10C2)"  ' R1C1 引用样式
Debug.Print Range("A2").Value       ' 取计算结果
```

> 在 VBA 里写公式必须用**英文函数名和逗号**，即使你的 Excel 是中文版用分号。

### Font：字体

```vb
With Range("A1").Font
    .Name = "微软雅黑"
    .Size = 14
    .Bold = True
    .Color = RGB(255, 0, 0)      ' 红色
    .Italic = True
End With
```

### Interior：填充

```vb
With Range("A1").Interior
    .Color = RGB(255, 255, 0)    ' 黄色背景
    .Pattern = xlSolid           ' 实心填充
End With
```

### Borders：边框

```vb
With Range("A1:C3").Borders
    .LineStyle = xlContinuous    ' 实线
    .Weight = xlThin             ' 细线
    .Color = RGB(0, 0, 0)
End With

' 只要外边框
Range("A1:C3").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
```

### 其他常用属性

```vb
Range("A1").NumberFormat = "yyyy-mm-dd"   ' 数字格式
Range("A1").NumberFormat = "0.00%"        ' 百分比
Range("A1:C3").HorizontalAlignment = xlCenter  ' 居中
Range("A1").WrapText = True               ' 自动换行
Range("A1").MergeCells                    ' 是否合并（只读判断）
Range("A1:B2").Merge                      ' 合并
Range("A1").RowHeight = 30
Range("B1").ColumnWidth = 20
```

## 7. 常用方法

```vb
Sub MethodDemo()
    ' 复制粘贴
    Range("A1:A5").Copy                    ' 复制
    Range("C1").PasteSpecial xlPasteValues ' 只粘贴值
    Application.CutCopyMode = False        ' 清除复制虚线框

    ' 删除 / 插入（带 shifting 参数）
    Range("A2").Delete Shift:=xlUp         ' 删除后下方上移
    Range("A2").Insert Shift:=xlDown       ' 插入后下方下移

    ' 清除
    Range("A1:C10").Clear                 ' 内容+格式全清
    Range("A1:C10").ClearContents         ' 只清内容
    Range("A1:C10").ClearFormats          ' 只清格式

    ' 排序
    Range("A1:C100").Sort _
        Key1:=Range("B1"), Order1:=xlDescending, Header:=xlYes

    ' 自动筛选
    Range("A1:C100").AutoFilter Field:=2, Criteria1:="北京"

    ' 查找
    Dim found As Range
    Set found = Range("A:A").Find(What:="张三", LookIn:=xlValues)
    If Not found Is Nothing Then MsgBox "在" & found.Address & "找到"

    ' 替换
    Range("A:A").Replace What:="旧", Replacement:="新"
End Sub
```

> `Find` 找不到时返回 `Nothing`，必须用 `If Not found Is Nothing` 判断，否则报错。

## 8. With 语句

对同一对象多次操作时，避免重复书写：

```vb
' 不用 With
Range("A1").Font.Bold = True
Range("A1").Font.Size = 14
Range("A1").Interior.Color = RGB(255, 255, 0)

' 用 With：简洁且更快
With Range("A1")
    .Font.Bold = True
    .Font.Size = 14
    .Interior.Color = RGB(255, 255, 0)
    .Value = "标题"
End With
```

With 可以嵌套，内层 `.Font` 指的是外层对象的 Font。

## 9. 对象浏览器

按 `F2` 打开对象浏览器，可以查看所有对象、属性、方法的说明。选中一个成员，底部会显示语法和简要说明，是查资料最快的方式。

使用技巧：在代码中选中一个关键字（如 `Range`），按 `Shift + F2` 直接跳转到它的定义。

## 10. 事件

事件是"发生某事时自动执行的代码"，写在特定位置：

### Workbook_Open：打开工作簿时执行

双击 `ThisWorkbook`，在代码窗口顶部下拉框选 `Workbook`，会自动生成：

```vb
Private Sub Workbook_Open()
    MsgBox "欢迎使用报表系统！"
    ' 也可以在这里做初始化：刷新数据、检查版本等
End Sub
```

### Worksheet_Change：单元格变化时执行

右键工作表标签 → 查看代码，粘贴：

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Target 是被修改的单元格
    If Target.Column = 1 Then   ' 只监控A列
        Application.EnableEvents = False   ' 防止递归触发
        Target.Offset(0, 1).Value = Now     ' B列记录修改时间
        Application.EnableEvents = True
    End If
End Sub
```

> **重要**：在 Change 事件里修改单元格会再次触发 Change 事件，形成无限递归。必须用 `Application.EnableEvents = False/True` 包裹。

### 常用事件一览

| 位置 | 事件 | 触发时机 |
|------|------|----------|
| ThisWorkbook | `Workbook_Open` | 打开工作簿 |
| ThisWorkbook | `Workbook_BeforeClose` | 关闭工作簿前（可取消关闭） |
| ThisWorkbook | `Workbook_SheetChange` | 任意工作表内容变化 |
| 工作表模块 | `Worksheet_Change` | 本表内容变化 |
| 工作表模块 | `Worksheet_SelectionChange` | 选中区域变化 |
| 工作表模块 | `Worksheet_BeforeDoubleClick` | 双击单元格前 |

事件过程必须是 `Private Sub`，且**不能改名**，否则不会被触发。
