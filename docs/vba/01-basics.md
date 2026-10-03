# VBA 基础

本章带你从零开始认识 VBA：打开编辑器、声明变量、使用运算符、写条件和循环、操作数组与字符串，并理解 Sub 过程与 Function 函数的区别。

## 1. 打开 VBE 编辑器

VBE（Visual Basic Editor）是编写 VBA 代码的地方，有三种打开方式：

- 快捷键：`Alt + F11`
- 功能区：「开发工具」选项卡 → 「Visual Basic」（如果没有该选项卡，到「文件 → 选项 → 自定义功能区」中勾选「开发工具」）
- 右键工作表标签 → 「查看代码」（直接进入该工作表对应的代码窗口）

### 插入模块并运行宏

1. 在 VBE 中点击菜单「插入 → 模块」，得到一个空白代码窗口（Module1）。
2. 输入代码，例如：

```vb
Sub HelloWorld()
    MsgBox "你好，VBA！"
End Sub
```

3. 把光标放在过程中，按 `F5` 运行；或点击工具栏的绿色三角 ▶。
4. 也可以回到 Excel 按 `Alt + F8`，在宏列表中选择并运行。

> **提示**：保存含宏的工作簿时要用 `.xlsm` 格式，`.xlsx` 会丢掉所有代码。

## 2. 变量声明与数据类型

### Dim 声明变量

```vb
Sub DeclareDemo()
    Dim name As String      ' 声明字符串变量
    Dim age As Integer      ' 声明整数变量
    Dim price As Double     ' 声明双精度浮点数

    name = "张三"
    age = 30
    price = 99.9

    MsgBox name & "今年" & age & "岁"
End Sub
```

### 常用数据类型

| 类型 | 说明 | 示例 |
|------|------|------|
| `String` | 字符串 | `"Hello"` |
| `Integer` | 整数（-32768 ~ 32767） | `100` |
| `Long` | 长整数（约 ±21 亿） | `100000` |
| `Single` | 单精度小数 | `3.14` |
| `Double` | 双精度小数 | `3.14159265` |
| `Boolean` | 布尔值 | `True` / `False` |
| `Date` | 日期时间 | `#2026-10-03#` |
| `Variant` | 万能类型（可存任何值） | 任意 |

> **建议**：整数优先用 `Long` 而不是 `Integer`。现代计算机上 `Long` 更快，且行号超过 32767 时 `Integer` 会溢出报错。

### 类型声明符（简写）

```vb
Dim s$      ' 等价于 Dim s As String
Dim n%      ' 等价于 Dim n As Integer
Dim l&      ' 等价于 Dim l As Long
Dim d#      ' 等价于 Dim d As Double
```

### Option Explicit：强制声明

在模块顶部写 `Option Explicit`，则所有变量必须先 `Dim` 才能使用。拼写错误会被立即发现，而不是默默产生一个新变量。

```vb
Option Explicit

Sub Test()
    Dim count As Long
    count = 10
    ' cont = 5   ' 拼错变量名会直接报错，而不是创建新变量
End Sub
```

> **强烈建议**：在 VBE「工具 → 选项 → 编辑器」中勾选「要求变量声明」，以后新建模块会自动加上 `Option Explicit`。

## 3. 运算符

### 算术运算符

```vb
Dim a As Long, b As Long
a = 10: b = 3
Debug.Print a + b   ' 13  加
Debug.Print a - b   ' 7   减
Debug.Print a * b   ' 30  乘
Debug.Print a / b   ' 3.333  除
Debug.Print a \ b   ' 3   整除（丢掉小数）
Debug.Print a Mod b ' 1   取余数
Debug.Print a ^ b   ' 1000 乘方
```

### 比较运算符

`=`, `<>`（不等于）, `>`, `<`, `>=`, `<=`。注意字符串比较用 `Like` 可做通配符匹配：

```vb
If "ABC123" Like "ABC*" Then MsgBox "以ABC开头"   ' * 匹配任意多个字符
If "A1" Like "A#" Then MsgBox "A后面跟一个数字"    ' # 匹配一个数字
If "AB" Like "A?" Then MsgBox "A后面跟一个字符"    ' ? 匹配一个字符
```

### 逻辑运算符

`And`（与）、`Or`（或）、`Not`（非）、`Xor`（异或）。

### 连接运算符

字符串连接用 `&`（推荐），`+` 也可以但遇到数字字符串时可能做加法，`&` 更安全：

```vb
Dim s As String
s = "总分：" & 100        ' "总分：100"
```

## 4. 条件语句

### If...Then...Else

```vb
Sub Grade(score As Long)
    If score >= 90 Then
        MsgBox "优秀"
    ElseIf score >= 60 Then
        MsgBox "及格"
    Else
        MsgBox "不及格"
    End If
End Sub
```

### Select Case

条件分支较多时比一串 `ElseIf` 更清晰：

```vb
Sub WeekdayName(n As Long)
    Select Case n
        Case 1
            MsgBox "星期一"
        Case 2 To 5
            MsgBox "工作日"
        Case 6, 7
            MsgBox "周末"
        Case Else
            MsgBox "输入错误"
    End Select
End Sub
```

`Case` 支持多种写法：单个值（`Case 1`）、范围（`Case 2 To 5`）、多个值（`Case 6, 7`）、比较（`Case Is > 100`）。

### IIf：单行条件函数

```vb
Dim result As String
result = IIf(score >= 60, "及格", "不及格")   ' 注意：两边都会被求值
```

> **注意**：`IIf` 是函数，两个分支都会执行，不适合放有副作用的代码。

## 5. 循环

### For...Next：固定次数

```vb
Sub SumDemo()
    Dim i As Long, total As Long
    For i = 1 To 100
        total = total + i
    Next i
    MsgBox "1到100的和是：" & total   ' 5050
End Sub

' 倒序 + 步长
Sub ReverseLoop()
    Dim i As Long
    For i = 10 To 1 Step -2
        Debug.Print i   ' 10, 8, 6, 4, 2
    Next i
End Sub
```

### For Each：遍历集合

```vb
Sub LoopSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print ws.Name
    Next ws
End Sub
```

### Do While / Do Until

```vb
Sub DoWhileDemo()
    Dim i As Long
    i = 1
    Do While i <= 5
        Debug.Print i
        i = i + 1
    Loop
End Sub

Sub DoUntilDemo()
    Dim i As Long
    i = 1
    Do Until i > 5      ' 条件成立时退出
        Debug.Print i
        i = i + 1
    Loop
End Sub
```

### Exit 跳出循环

```vb
Sub FindDemo()
    Dim i As Long
    For i = 1 To 1000
        If Cells(i, 1).Value = "目标" Then
            MsgBox "在第" & i & "行找到"
            Exit For    ' 跳出 For 循环
        End If
    Next i
End Sub
```

`Exit For` 跳出 For 循环，`Exit Do` 跳出 Do 循环。

## 6. 数组

### 静态数组

```vb
Sub StaticArray()
    Dim arr(1 To 5) As String   ' 1到5，共5个元素
    Dim i As Long

    arr(1) = "一月": arr(2) = "二月": arr(3) = "三月"
    arr(4) = "四月": arr(5) = "五月"

    For i = 1 To 5
        Debug.Print arr(i)
    Next i
End Sub
```

### 动态数组

大小运行时决定，用 `ReDim` 调整：

```vb
Sub DynamicArray()
    Dim arr() As Long
    Dim n As Long
    n = 10

    ReDim arr(1 To n)          ' 分配10个元素
    ' ... 使用 ...
    ReDim Preserve arr(1 To n + 5)  ' 扩大到15个，Preserve 保留原有数据
End Sub
```

> 不加 `Preserve` 时 `ReDim` 会清空数组。

### 一维数组快速赋值：Array 函数

```vb
Dim week As Variant
week = Array("一", "二", "三", "四", "五", "六", "日")
Debug.Print week(0)   ' 注意：Array 返回的数组下标从 0 开始
```

### 二维数组：批量读写单元格（性能关键）

逐个单元格读写很慢，一次读入数组、内存中处理、一次写回，快 10 倍以上：

```vb
Sub FastReadWrite()
    Dim data As Variant
    Dim i As Long

    ' 一次读入 A1:C1000 到二维数组（下标从1开始）
    data = Range("A1:C1000").Value

    For i = 1 To UBound(data, 1)       ' UBound(data,1) = 行数
        data(i, 3) = data(i, 1) * data(i, 2)
    Next i

    ' 一次写回
    Range("A1:C1000").Value = data
End Sub
```

## 7. 字符串处理

```vb
Sub StringDemo()
    Dim s As String
    s = "  Hello, VBA  "

    Debug.Print Trim(s)              ' "Hello, VBA"  去两端空格
    Debug.Print LCase(s)             ' 全小写
    Debug.Print UCase(s)             ' 全大写
    Debug.Print Len("你好")          ' 2  字符数
    Debug.Print Left("ABCDEF", 3)    ' "ABC"  取左边
    Debug.Print Right("ABCDEF", 2)   ' "EF"  取右边
    Debug.Print Mid("ABCDEF", 2, 3)  ' "BCD"  从第2个起取3个
    Debug.Print InStr("ABCDEF", "CD")' 3  查找位置（找不到返回0）
    Debug.Print Replace("a-b-c", "-", "+")  ' "a+b+c"  替换
    Debug.Print Split("a,b,c", ",")(1)      ' "b"  分割成数组（从0开始）
    Debug.Print Join(Array("a","b","c"), "-") ' "a-b-c"  连接数组
End Sub
```

## 8. Sub 过程与 Function 函数

| | Sub 过程 | Function 函数 |
|---|---|---|
| 作用 | 执行一系列操作 | 计算并**返回一个值** |
| 调用 | `Call MySub` 或直接 `MySub` | `x = MyFunc(1, 2)` 或在单元格中 `=MyFunc(A1)` |
| 返回值 | 无 | 有（函数名 = 返回值） |

```vb
' Sub：做事
Sub Greet(name As String)
    MsgBox "你好，" & name
End Sub

' Function：算值
Function AddTax(price As Double) As Double
    AddTax = price * 1.13   ' 函数名 = 返回值
End Function

Sub Test()
    Greet "张三"                 ' 调用 Sub
    MsgBox AddTax(100)           ' 调用 Function，结果 113
End Sub
```

**自定义函数（UDF）**：写在标准模块里的 `Function` 可以直接在工作表公式中使用，比如在单元格输入 `=AddTax(A1)`。

## 9. 注释与代码规范

```vb
' 单行注释用单引号

Sub Standard()
    Dim userName As String   ' 变量名：驼峰式，见名知意
    Dim totalAmount As Double

    userName = "张三"
    totalAmount = 1234.56

    ' 复杂逻辑前写注释说明意图，而不是复述代码
    ' 计算含税总价（税率13%）
    totalAmount = totalAmount * 1.13
End Sub
```

### 规范建议

1. **模块顶部写 `Option Explicit`**，强制声明变量。
2. **变量命名**：驼峰式（`userName`），过程名动词开头（`GetTotal`、`SaveReport`）。
3. **缩进**：每个嵌套层级缩进 4 个空格（Tab 键）。
4. **一行只做一件事**，过程尽量短（超过 50 行考虑拆分）。
5. **魔法数字**用常量代替：`Const TAX_RATE As Double = 0.13`。
6. 长语句用空格加下划线 ` _` 换行：

```vb
MsgBox "这是一条很长的消息，" _
     & "用下划线换行书写更清晰。"
```
