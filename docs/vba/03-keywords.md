# VBA 关键字大全

本章按类别系统列出 VBA 常用关键字的语法、示例和注意事项。建议收藏作为速查手册。

---

## 一、声明类

### Dim — 声明变量

```vb
Dim name As String          ' 声明单个变量
Dim x As Long, y As Long    ' 一行声明多个
Dim arr(1 To 10) As Long    ' 声明数组
Dim ws As Worksheet         ' 声明对象变量
```

作用域：过程内部，只能在当前过程中使用。

### Private — 模块级私有声明

```vb
Private counter As Long     ' 模块顶部，整个模块可用，外部模块不可见

Private Sub Helper()
    ' 私有过程，只能被本模块调用
End Sub
```

### Public — 公开声明

```vb
Public appName As String    ' 整个工程可用

Public Sub Main()
    ' 公开过程，可被其他模块调用，也可出现在 Alt+F8 宏列表
End Sub
```

### Static — 静态变量（保持值）

普通局部变量在过程结束後就销毁；`Static` 声明的变量在多次调用之间**保留上次的值**：

```vb
Sub CountCalls()
    Static times As Long   ' 只初始化一次
    times = times + 1
    MsgBox "这是第 " & times & " 次调用"
End Sub
```

也可以让整个过程的变量都静态：`Static Sub CountCalls()`。

### Const — 常量

```vb
Const TAX_RATE As Double = 0.13
Const APP_NAME As String = "报表系统"
' Const MAX_ROWS As Long = 1048576
```

常量在编译时确定，不能在运行中修改。魔法数字（税率、上限等）都应该用常量。

### ReDim — 重定义动态数组

```vb
Dim arr() As Long
ReDim arr(1 To 10)              ' 分配空间
ReDim Preserve arr(1 To 20)     ' 扩大并保留原有数据
```

规则：
- 只能用于动态数组（声明时不写上下界）。
- `ReDim` 会清空数组，加 `Preserve` 则保留数据。
- `Preserve` 只能改变**最后一维**的大小。

### Option Explicit — 强制变量声明

```vb
Option Explicit   ' 必须写在模块最顶部，所有代码之前
```

开启后，未经 `Dim` 的变量直接报错。防止拼写错误产生难以察觉的 bug。**所有模块都应该加。**

---

## 二、流程控制类

### If / Then / Else / ElseIf

```vb
If 条件1 Then
    ' 条件1成立
ElseIf 条件2 Then
    ' 条件2成立
Else
    ' 都不成立
End If

' 单行写法（简单情况）
If x > 0 Then MsgBox "正数"
```

### Select / Case — 多分支选择

```vb
Select Case score
    Case Is >= 90
        grade = "A"
    Case 80 To 89
        grade = "B"
    Case 70, 71, 72
        grade = "C"
    Case Else
        grade = "D"
End Select
```

`Case` 的四种形式：`Is` 比较、`To` 范围、逗号多值、`Else` 兜底。

### For / Next — 计数循环

```vb
For i = 1 To 10 Step 2    ' Step 可省略（默认1），可为负数
    Debug.Print i
Next i
```

### For Each — 集合遍历

```vb
Dim ws As Worksheet
For Each ws In Worksheets
    Debug.Print ws.Name
Next ws
```

注意：`For Each` 的循环变量**必须是 Variant 或对象类型**，不能是 `Long`。

### Do / Loop — 条件循环

四种形式：

```vb
Do While 条件      ' 先判断，后执行（可能一次都不执行）
    ' ...
Loop

Do Until 条件      ' 条件成立时退出
    ' ...
Loop

Do                 ' 先执行，后判断（至少执行一次）
    ' ...
Loop While 条件

Do
    ' ...
Loop Until 条件
```

### While / Wend — 旧式循环

```vb
While i < 10
    i = i + 1
Wend
```

功能与 `Do While...Loop` 相同，是老式写法。新代码推荐用 `Do...Loop`。

### GoTo — 跳转

```vb
Sub Demo()
    On Error GoTo ErrHandler
    ' ... 正常代码 ...
    Exit Sub

ErrHandler:                  ' 标签以冒号结尾
    MsgBox "出错了：" & Err.Description
End Sub
```

现代 VBA 中 `GoTo` 主要用于错误处理。随意跳转会让代码难以维护，**不要**用它代替正常的条件分支。

### Exit — 提前退出

```vb
Exit Sub        ' 退出 Sub 过程
Exit Function   ' 退出 Function 函数
Exit For        ' 跳出 For 循环
Exit Do         ' 跳出 Do 循环
```

常用于"卫语句"：先处理异常情况并退出，主逻辑保持平直：

```vb
Sub Process(ws As Worksheet)
    If ws Is Nothing Then
        MsgBox "工作表不存在"
        Exit Sub
    End If
    ' 主逻辑...
End Sub
```

### End — 终止程序

```vb
End   ' 立即终止所有代码执行，释放所有变量
```

> **警告**：`End` 会粗暴终止一切，不触发任何清理代码。正常结束过程用 `Exit Sub`，几乎用不到 `End`。

---

## 三、过程定义类

### Sub — 过程（做事）

```vb
Sub SayHello()
    MsgBox "你好"
End Sub

Sub Greet(name As String, Optional title As String = "先生")
    MsgBox title & name & "，你好"
End Sub
```

### Function — 函数（算值并返回）

```vb
Function CircleArea(r As Double) As Double
    CircleArea = 3.14159 * r * r   ' 函数名 = 返回值
End Function

' 提前返回
Function SafeDiv(a As Double, b As Double) As Variant
    If b = 0 Then
        SafeDiv = "除数不能为零"
        Exit Function
    End If
    SafeDiv = a / b
End Function
```

### Call — 调用过程

```vb
Call Greet("张三")   ' 显式调用，加括号
Greet "张三"         ' 省略 Call，不加括号（效果相同）
```

> 省略 `Call` 是更常见的写法。注意：省略 Call 时参数不能加括号，否则会被当作求值。

### ByVal / ByRef — 传值 vs 传址

```vb
Sub ByValDemo(ByVal x As Long)   ' 传值：复制一份，内部修改不影响外部
    x = x + 10
End Sub

Sub ByRefDemo(ByRef x As Long)   ' 传址：操作原变量，内部修改影响外部
    x = x + 10
End Sub

Sub Test()
    Dim n As Long
    n = 5
    ByValDemo n
    Debug.Print n   ' 还是 5
    ByRefDemo n
    Debug.Print n   ' 变成 15
End Sub
```

**VBA 默认是 `ByRef`**（与大多数语言相反）。如果不希望过程修改传入的变量，显式写 `ByVal`。

### Optional — 可选参数

```vb
Sub Print2D(rng As Range, Optional copies As Long = 1)
    Dim i As Long
    For i = 1 To copies
        rng.PrintOut
    Next i
End Sub

Print2D Range("A1")          ' copies 用默认值 1
Print2D Range("A1"), 3       ' 打印 3 份
```

规则：`Optional` 参数必须有默认值，且必须放在参数列表最后。

### ParamArray — 可变数量参数

```vb
Function SumAll(ParamArray nums() As Variant) As Double
    Dim n As Variant, total As Double
    For Each n In nums
        total = total + n
    Next n
    SumAll = total
End Function

Debug.Print SumAll(1, 2, 3)          ' 6
Debug.Print SumAll(10, 20, 30, 40)   ' 100
```

`ParamArray` 只能有一个，必须是最后一个参数，类型必须是 `Variant` 数组。

---

## 四、对象操作类

### Set — 对象赋值

```vb
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("数据")   ' 对象必须用 Set

Dim n As Long
n = 100                                    ' 普通变量直接 =
```

忘记 `Set` 会报运行时错误 424"需要对象"。

### New — 创建对象实例

```vb
Dim dict As Object
Set dict = New Scripting.Dictionary   ' 需要引用 Microsoft Scripting Runtime

' 或者声明时直接 New（不推荐：每次使用都检查是否已创建，有开销）
Dim dict2 As New Scripting.Dictionary
```

### Nothing — 空对象 / 释放对象

```vb
Dim ws As Worksheet
Set ws = Nothing   ' 释放对象引用

' 判断对象是否为空
If ws Is Nothing Then
    MsgBox "对象未初始化"
End If
```

对象比较必须用 `Is` / `Is Not`，不能用 `=`。

### With — 对象简化书写

```vb
With Range("A1")
    .Value = "标题"
    .Font.Bold = True
    .Font.Size = 16
    .Interior.Color = RGB(255, 255, 0)
End With
```

### Me — 当前对象自身

在**工作表/工作簿模块**或**用户窗体**代码中，`Me` 指代当前对象：

```vb
' 在 Sheet1 的代码模块中
Private Sub Worksheet_Change(ByVal Target As Range)
    Me.Range("B1").Value = Now   ' Me = Sheet1 本身
End Sub

' 在用户窗体中
Private Sub UserForm_Click()
    Me.Caption = "被点了"        ' Me = 这个窗体
End Sub
```

---

## 五、错误处理类

### On Error — 错误处理开关

三种模式：

```vb
' 1. 出错时跳转到标签
Sub Demo1()
    On Error GoTo ErrHandler
    ' ... 可能出错的代码 ...
    Exit Sub
ErrHandler:
    MsgBox "出错：" & Err.Description
End Sub

' 2. 出错时继续执行下一行（用于可预期的、要逐个处理的情况）
Sub Demo2()
    On Error Resume Next
    Workbooks("不存在.xlsx").Close   ' 如果文件没打开也不报错
    On Error GoTo 0                  ' 恢复默认（报错即中断）
End Sub

' 3. 出错时跳转到下一行（Resume Next 的标签版）
Sub Demo3()
    On Error GoTo SkipIt
    ' ...
SkipIt:
    Resume Next
End Sub
```

> `On Error Resume Next` 是双刃剑：它会**吞掉所有错误**。使用后必须立即检查 `Err.Number`，并尽快用 `On Error GoTo 0` 恢复。

### Resume — 恢复执行

```vb
Sub Demo()
    On Error GoTo Handler
    x = 1 / 0
    Exit Sub
Handler:
    MsgBox "除数不能为零"
    Resume Next        ' 从出错的下一行继续
    ' Resume           ' 从出错的那一行重试
    ' Resume ExitHere   ' 跳转到指定标签
ExitHere:
End Sub
```

### Err — 错误对象

```vb
Sub ErrDemo()
    On Error Resume Next
    Dim n As Long
    n = "abc"                  ' 类型不匹配错误

    If Err.Number <> 0 Then
        Debug.Print "错误号：" & Err.Number         ' 13
        Debug.Print "错误描述：" & Err.Description  ' 类型不匹配
        Debug.Print "出错过程：" & Err.Source
        Err.Clear                  ' 清除错误状态
    End If
    On Error GoTo 0
End Sub
```

常见错误号：13（类型不匹配）、9（下标越界）、424（需要对象）、91（对象变量未设置）、1004（Excel 内部错误）。

---

## 六、逻辑与其他

### And / Or / Not — 逻辑运算

```vb
If age >= 18 And age <= 60 Then MsgBox "劳动年龄"
If city = "北京" Or city = "上海" Then MsgBox "一线城市"
If Not done Then MsgBox "还没完成"
```

注意 VBA 没有短路求值：`If A And B` 中 A、B 都会被计算。

### True / False — 布尔字面量

```vb
Dim done As Boolean
done = True
If done = False Then   ' 也可以写 If Not done Then
```

### Empty / Null / Nothing 的区别

这是新手最容易混淆的三个概念：

| 关键字 | 含义 | 出现场景 |
|--------|------|----------|
| `Empty` | **未初始化的 Variant** | `Dim v As Variant` 后，`v` 就是 Empty |
| `Null` | **无效数据**（数据库概念） | 从数据库读到空字段；`Null` 参与任何运算结果都是 `Null` |
| `Nothing` | **空对象引用** | 对象变量 `Set` 之前；`Find` 找不到时返回 |

```vb
Sub CompareDemo()
    Dim v As Variant
    Debug.Print IsEmpty(v)       ' True，Variant 初始状态

    Dim ws As Worksheet
    ' Set ws = ...  ' 没执行
    Debug.Print ws Is Nothing    ' True

    Dim x As Variant
    x = Null
    Debug.Print IsNull(x)        ' True
    Debug.Print IsNull(x + 1)    ' True！Null 参与运算结果还是 Null

    ' 判断单元格是否为空
    Debug.Print IsEmpty(Range("A1").Value)   ' A1没内容时 True
End Sub
```

配套判断函数：`IsEmpty()`、`IsNull()`、`Is Nothing`（注意 `Is` 前后空格）。

### 其他常用关键字速览

| 关键字 | 用途 | 示例 |
|--------|------|------|
| `As` | 声明类型 | `Dim x As Long` |
| `To` | 范围 | `For i = 1 To 10` / `Case 1 To 5` |
| `Step` | 步长 | `For i = 10 To 1 Step -1` |
| `Each` | 遍历 | `For Each ws In Worksheets` |
| `Then` | 条件 | `If x > 0 Then` |
| `ElseIf` | 多条件 | 见 If 示例 |
| `Case` | 分支 | `Select Case` |
| `Loop` | 循环尾 | `Do While ... Loop` |
| `Wend` | 循环尾 | `While ... Wend` |
| `Is` | 对象比较 | `If ws Is Nothing` |
| `Like` | 通配符匹配 | `If s Like "A*"` |
| `Mod` | 取余 | `10 Mod 3 = 1` |
| `Type` | 自定义类型 | 见下 |
| `Enum` | 枚举 | 见下 |

### Type — 自定义数据类型

```vb
' 模块顶部定义
Type Employee
    Name As String
    Age As Long
    Salary As Double
End Type

Sub UseType()
    Dim emp As Employee
    emp.Name = "张三"
    emp.Age = 30
    emp.Salary = 8000
    MsgBox emp.Name & "的工资是" & emp.Salary
End Sub
```

### Enum — 枚举

给一组相关的常量起名字，提高可读性：

```vb
' 模块顶部定义
Enum WeekDay
    星期一 = 1
    星期二 = 2
    星期三 = 3
End Enum

Sub UseEnum()
    Dim d As WeekDay
    d = 星期三
    MsgBox d   ' 显示 3
End Sub
```
