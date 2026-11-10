# VBA 进阶：用户窗体、类模块、加载宏与错误处理进阶

本篇是 VBA 自动化（C11）的进阶篇：用用户窗体（UserForm）做交互界面，用类模块把数据和行为封装成对象，把常用工具打包成加载宏（.xlam）随 Excel 启动，用 Err 对象加集中处理与日志让工具在别人电脑上出错时也能定位原因。

- **所属栏目**：C11 VBA 自动化 ／ C11-L2-5 VBA 进阶
- **适用版本**：Excel 2010 / 2013 / 2016 / 2019 / 2021 / 365（Windows 桌面版；Mac 版 VBA 不支持 UserForm 与部分加载宏功能）
- **难度**：高级
- **预计阅读时长**：60 分钟
- **更新日期**：2026-11-10

> **验证说明**：本篇 VBA 语法经 Microsoft 官方文档核对，数值经 Python 独立验算；**所有 VBA 代码均未在 Excel 内执行验证**（无 Excel 环境）。请在副本工作簿上先行试跑，确认无误后再用于正式数据。

## 学习目标

学完本篇，你将能够：
- 用 UserForm 搭建带文本框、下拉框与按钮的录入窗体，并区分模态与非模态
- 用类模块封装数据与行为（如员工对象），并用集合管理多个对象
- 把工具打包成 .xlam 加载宏并安装到 Excel，随任何工作簿启动
- 用 Err 对象写集中错误处理：分错误号处理、记日志、恢复现场

## 1. 基础概念

### 1.1 用户窗体（UserForm）

UserForm 是 VBA 自带的对话框设计器：在 VBE 中「插入 → 用户窗体」，拖入文本框（TextBox）、标签（Label）、下拉框（ComboBox）、按钮（CommandButton）等控件，双击控件写事件代码。窗体本身是一个对象，有 `Show`、`Hide`、`Unload` 三个基本动作：

- `Show`：显示窗体。默认模态（vbModal，值为 1）：用户必须先关闭窗体才能操作工作表，`Show` 之后的代码暂停执行，直到窗体关闭或隐藏。
- `Show vbModeless`（值为 0）：非模态：用户可以一边开着窗体一边操作工作表，`Show` 之后的代码立即继续执行。
- `Hide`：隐藏窗体但保留其中的输入值；`Unload Me`：彻底卸载，输入值丢失。

窗体有两个最常用的事件：`UserForm_Initialize`（窗体第一次加载时执行一次，适合填下拉框初始值）和各控件的 `Click` / `Change` 事件。窗体代码写在窗体自己的代码窗口里（双击窗体打开），而「显示窗体」的那行 `frmEntry.Show` 通常写在标准模块里。

### 1.2 类模块（Class Module）

类模块是自定义对象的蓝图：在 VBE 中「插入 → 类模块」，把一类事物的数据（属性）与行为（方法）写在一起。标准模块里的变量是散的，类模块把它们收拢：

- 属性用 `Property Get`（读）与 `Property Let`（写）定义，写在属性过程里可以顺带做校验（例如工资不能为负）。
- 方法就是写在类里的 `Sub` / `Function`。
- 使用时先 `Dim emp As New clsEmployee`（或 `Set emp = New clsEmployee`），再 `emp.Name = "张三"`。

类模块的文件扩展名是 .cls，标准模块是 .bas；两者导入方式相同（VBE「文件 → 导入文件」），不要把类代码粘贴进标准模块，否则 `Property Get` 等写法会编译报错。

### 1.3 加载宏（.xlam）

加载宏是「保存为 .xlam 格式、启动时自动加载的工作簿」：代码随 Excel 启动而可用，工作表界面被隐藏，函数与宏可以在任何打开的工作簿里调用。制作步骤是先在普通工作簿里写好并调通代码，再「文件 → 另存为 → 保存类型选 Excel 加载宏 (*.xlam)」，Excel 会默认保存到加载宏目录；然后「文件 → 选项 → 加载项 → 管理选 Excel 加载项 → 转到」，勾选你的加载宏完成安装。安装后加载宏每次启动 Excel 自动加载，不用再打开源文件。

与 .xlsm 的区别：.xlsm 是「带宏的工作簿」，代码只在该文件打开时可用；.xlam 是「随 Excel 启动的工具库」，代码对所有工作簿可用，但它的工作表不可见，不适合存放业务数据。调试期保留一份 .xlsm 源文件，改好后再另存为 .xlam（直接改 .xlam 里的代码容易因保存路径问题丢改动）。

### 1.4 错误处理进阶：Err 对象与集中处理

VBA 运行出错时，内建的 `Err` 对象会被填入当次错误的信息（Microsoft Learn 的 Err 对象文档）：

| 成员 | 说明 |
|------|------|
| `Err.Number` | 错误号，如 13（类型不匹配）、9（下标越界）、424（需要对象）、91（对象变量未设置）、1004（Excel 内部错误） |
| `Err.Description` | 错误描述文字，如"类型不匹配" |
| `Err.Source` | 出错的源（过程名或对象名，可自行写入） |
| `Err.Clear` | 清空错误状态（Number 归 0） |
| `Err.Raise` | 主动抛出一个错误；系统错误号保留 0–512，自定义错误号用 512 以上 |

进阶做法是「集中处理」：每个业务过程只写一行 `On Error GoTo ErrHandler`，出错后统一跳到文末的错误处理段；处理段按 `Err.Number` 分情况（可恢复的记日志继续、不可恢复的提示后退出），并把「时间、过程名、错误号、错误描述」写进日志工作表或文本文件。日志是交付工具后唯一的排障依据：用户只说"报错了"没法定位，有日志才能看到是哪一行、哪个错误号。

`On Error Resume Next` 只适合「可预期的单个操作」（如判断某工作表是否存在、关闭一个可能没打开的文件），用完必须立即检查 `Err.Number` 并用 `On Error GoTo 0` 恢复；把它盖在整段业务代码上会吞掉所有错误，是最常见的坏习惯。

## 2. 操作步骤

### 2.1 做一个录入窗体（UserForm）

1. 在 VBE 中点击「插入 → 用户窗体」，在属性窗口（F4）把窗体名称改为 `frmEntry`、标题（Caption）改为"报销录入"。
2. 从工具箱拖入控件：2 个标签（姓名、金额）、1 个文本框（txtName）、1 个文本框（txtAmount）、1 个下拉框（cmbType）、2 个按钮（cmdOK 标题"确定"、cmdCancel 标题"取消"）。
3. 双击窗体空白处写 `UserForm_Initialize`：往 cmbType 里加"差旅"、"办公"、"招待"三项，默认选中第一项。
4. 双击 cmdOK 写 `cmdOK_Click`：校验姓名非空、金额是数字且大于 0；校验通过后把一行数据写到"报销"表末尾，再 `Unload Me`。
5. 双击 cmdCancel 写 `cmdCancel_Click`：只执行 `Unload Me`，不写数据。
6. 在标准模块写 `ShowEntryForm`：`frmEntry.Show`（模态，录入完成前不让用户碰工作表）。
7. 按 F5 运行 `ShowEntryForm`，在窗体里输一行数据，检查"报销"表是否多了一行。

### 2.2 写一个类模块

1. 在 VBE 中点击「插入 → 类模块」，在属性窗口把名称改为 `clsEmployee`。
2. 在类模块顶部声明私有字段：`Private pName As String`、`Private pSalary As Double`。
3. 为每个字段写一对 `Property Get` / `Property Let`；在 `Salary` 的 `Let` 里加校验：小于 0 时用 `Err.Raise 513, , "工资不能为负数"` 拒绝写入。
4. 写一个方法 `AnnualPay`：返回 `pSalary * 12`（年薪，不含奖金）。
5. 在标准模块写测试过程：建 3 个 `clsEmployee` 对象放进 `Collection`，循环累加 `Salary` 求合计与平均。
6. 按 F5 运行测试过程，用立即窗口（Ctrl+G）核对合计与平均值。

### 2.3 打包成加载宏（.xlam）

1. 新建一个工作簿，把要共享的过程（如本篇第 4 节的错误处理模块、第 6 节的工具过程）放进标准模块，先在 .xlsm 里调通。
2. 在 VBE「调试 → 编译」确认无编译错误；把 `ThisWorkbook` 的 `IsAddin` 属性记在心里（另存为 .xlam 后它自动变为 True，工作表随之隐藏）。
3. 回到 Excel，「文件 → 另存为 → 保存类型选 Excel 加载宏 (*.xlam)」，起个名字（如 MyTools.xlam），保存在默认加载宏目录。
4. 「文件 → 选项 → 加载项 → 管理选 Excel 加载项 → 转到」，勾选 MyTools，点确定。
5. 关闭并重新打开 Excel，确认加载宏已自动加载（VBE 工程窗口里能看到它）。
6. 新建一个空白工作簿，运行加载宏里的任一公开过程，确认跨工作簿可用。

### 2.4 接入集中错误处理与日志

1. 新建标准模块 `modErrorHandler`，写 `LogError`（把时间、过程名、错误号、错误描述追加写到同目录 `ErrorLog.txt`）与 `HandleError`（按错误号分支：9/13/1004 提示用户检查数据，其余提示联系管理员）。
2. 在每个业务过程顶部加 `On Error GoTo ErrHandler`，正常出口前加 `Exit Sub`，文末加 `ErrHandler:` 标签段，段内调 `HandleError "过程名"`。
3. 在 `ErrHandler` 段内恢复现场：`Application.ScreenUpdating = True`、`Application.EnableEvents = True`（业务过程里关掉的开关必须在这里开回来，否则出错一次后 Excel 就"假死"了）。
4. 故意制造一次错误（如让代码打开不存在的工作表），运行后打开 `ErrorLog.txt`，确认日志里有一行带正确错误号的记录。

## 3. 图文示例

### 示例 1：报销录入窗体（控件一览）

窗体 `frmEntry`（待补图）：标题"报销录入"；第一行标签"姓名"加文本框 txtName；第二行标签"金额"加文本框 txtAmount；第三行标签"类型"加下拉框 cmbType（三项：差旅、办公、招待）；底部两个按钮 cmdOK（确定）、cmdCancel（取消）。

文字示例：用户打开窗体 → 姓名输"张三"、金额输"1200"、类型选"差旅" → 点确定 → "报销"表末尾新增一行（张三、1200、差旅），窗体关闭；金额输"abc"点确定 → 弹出"金额必须是大于 0 的数字"，窗体不关闭、不写数据。

### 示例 2：员工类与集合（3 名员工）

| 姓名 | 月薪 | 年薪（月薪 × 12） |
|------|------|-------------------|
| 张三 | 8000 | 96000 |
| 李四 | 9500 | 114000 |
| 王五 | 12000 | 144000 |
| 合计 | 29500 | 354000 |
| 平均月薪 | 9833.33 | — |

运行结果：合计月薪 → 29500；平均月薪 → 9833.33（29500 ÷ 3，四舍五入到 2 位小数）；王五的年薪 → 144000。

### 示例 3：集中错误处理跑一次

故意让代码打开不存在的工作簿 `Workbooks.Open("D:\不存在.xlsx")`，错误号 → 1004；`HandleError` 提示"文件不存在或路径错误，请检查路径后重试"，并在 `ErrorLog.txt` 里记一行：`2026-11-10 09:00:00 | OpenReport | 1004 | 文件不存在或路径错误`（时间以实际运行为准）。

## 4. 公式和代码

### 4.1 UserForm：录入窗体代码

```vb
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
```

运行结果：姓名"张三"、金额 1200、类型"差旅" → "报销"表末尾新增一行（张三、1200、差旅、当天日期）；金额输"abc" → 弹出"金额必须是数字"，不写数据。

```vb
' ---- 以下代码放在标准模块（如 modMain），用来显示窗体 ----
Option Explicit

Sub ShowEntryForm()
    frmEntry.Show   ' 默认模态（vbModal）：关闭窗体前不能操作工作表
    ' frmEntry.Show vbModeless   ' 改成这一行则为非模态：可一边开窗体一边操作表格
End Sub
```

运行结果：`frmEntry.Show` → 弹出模态窗体，`Show` 之后的代码暂停，直到窗体关闭；改成 `vbModeless` 后代码立即继续执行。

### 4.2 类模块：员工类

```vb
' ---- 以下代码放在类模块 clsEmployee（插入 → 类模块，名称改为 clsEmployee）----
' ---- 注意：类模块代码不能粘贴进标准模块，否则编译报错 ----
Option Explicit

Private pName As String
Private pSalary As Double

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(ByVal v As String)
    pName = Trim(v)
End Property

Public Property Get Salary() As Double
    Salary = pSalary
End Property

Public Property Let Salary(ByVal v As Double)
    If v < 0 Then
        Err.Raise 513, "clsEmployee", "工资不能为负数"   ' 自定义错误号用 512 以上
    End If
    pSalary = v
End Property

Public Function AnnualPay() As Double
    AnnualPay = pSalary * 12
End Function
```

```vb
' ---- 以下代码放在标准模块，测试类模块 ----
Option Explicit

Sub TestEmployeeClass()
    Dim staff As Collection
    Dim emp As clsEmployee
    Dim total As Double
    Dim i As Long

    Set staff = New Collection

    Set emp = New clsEmployee
    emp.Name = "张三": emp.Salary = 8000
    staff.Add emp

    Set emp = New clsEmployee
    emp.Name = "李四": emp.Salary = 9500
    staff.Add emp

    Set emp = New clsEmployee
    emp.Name = "王五": emp.Salary = 12000
    staff.Add emp

    total = 0
    For i = 1 To staff.Count
        total = total + staff(i).Salary
        Debug.Print staff(i).Name & " 年薪 " & staff(i).AnnualPay()
    Next i

    Debug.Print "合计月薪 " & total              ' 29500
    Debug.Print "平均月薪 " & Round(total / staff.Count, 2)  ' 9833.33
End Sub
```

运行结果：立即窗口逐行输出"张三 年薪 96000"、"李四 年薪 114000"、"王五 年薪 144000"，最后两行 → 合计月薪 29500、平均月薪 9833.33。

### 4.3 加载宏：随启动可用的工具库

```vb
' ---- 以下代码放在标准模块 modAddIn，调通后随工作簿另存为 .xlam ----
Option Explicit

' 加载宏每次随 Excel 启动加载时自动执行一次（.xlam 里用 Auto_Open）
Sub Auto_Open()
    MsgBox "MyTools 加载宏已加载：工具在「加载宏」菜单中可用", vbInformation
End Sub

' 示例工具过程：给当前选中区域加外边框（任何工作簿里都能调）
Sub AddOuterBorder()
    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中一个单元格区域", vbExclamation
        Exit Sub
    End If
    Selection.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
End Sub
```

运行结果：安装加载宏后重启 Excel → 弹出"MyTools 加载宏已加载"；在任意工作簿选中区域运行 `AddOuterBorder` → 区域加上中等粗细外边框。

### 4.4 错误处理进阶：集中处理与日志

```vb
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
```

```vb
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
```

运行结果：文件不存在时 → 错误号 1004，弹出"文件不存在或路径错误"，`ErrorLog.txt` 增加一行（含时间、OpenReport、1004 与错误描述）。

## 5. 实际工作场景

### 场景 1：报销录入窗体（财务 / 行政）

每月几十张报销单，手工在表里逐行填容易串行、漏项。做法：按 2.1 节做 `frmEntry` 窗体，姓名、金额必填校验，类型只能从下拉框三选一，点确定才写一行，取消不写数据。结果：录入员不用记列顺序，金额输错当场拦截；`Unload Me` 保证下次打开是空窗体，不会把上个人的数据带进去。进阶：把窗体改成非模态（`Show vbModeless`），录入员可以一边对照原始单据表一边填，不用先关窗体。

### 场景 2：部门工具加载宏（全员共享）

部门有 5 个常用小工具（加边框、去空行、批量改名等），以前每人复制一份 .xlsm，各自版本越用越乱。做法：把 5 个过程收进一个工作簿，调通后另存为 MyTools.xlam，发给大家按 2.3 节安装；后续升级只发新版 .xlam 覆盖安装即可。结果：所有人打开任意工作簿都能用同一套工具；工具代码与业务数据彻底分离，业务表再也不会因为"顺手改了宏"而损坏。

## 6. 常见错误

| 错误现象 | 原因 | 解决方案 |
|---|---|---|
| 窗体弹出来是空白的，控件全不见了 | 控件画在了窗体外面，或窗体被意外改小 | 在 VBE 里点窗体，按 Ctrl+A 全选控件看有没有在边界外；把控件拖回窗体内 |
| `frmEntry.Show` 报"找不到工程或库" | 窗体被改名或删除，标准模块里的名字对不上 | 检查工程窗口里的窗体名与代码中的名字是否一致，改名后要同步改调用处 |
| 非模态窗体一闪而过 / 最小化 Excel 后窗体消失 | 非模态窗体没有用变量持有引用，或 Excel 最小化时窗体跟随隐藏 | 用 `Dim frm As New frmEntry: frm.Show vbModeless` 持有引用；或改回模态 |
| 类模块代码粘贴进标准模块后编译报错 | `Property Get/Let` 只能写在类模块里 | 插入真正的类模块（插入 → 类模块），把代码移过去 |
| `Salary = -100` 没有报错，写进去了 | 校验写在了 `Property Get` 里而不是 `Property Let` 里，或根本没写校验 | 校验必须写在 `Property Let` 里；用 `Err.Raise 513` 拒绝非法值 |
| 加载宏安装后找不到工具 | 只保存了 .xlam，没在「加载项 → 转到」里勾选 | 文件 → 选项 → 加载项 → 管理选 Excel 加载项 → 转到 → 勾选后确定 |
| 加载宏里改了代码，重启后改动丢了 | 直接在 .xlam 里改，保存时写回了加载宏目录以外的路径或根本没保存 | 保留 .xlsm 源文件，改源文件再重新另存为 .xlam 覆盖 |
| 出错一次后 Excel 按钮全变灰、事件不触发 | 错误处理段里没恢复 `ScreenUpdating` / `EnableEvents` | 在 `ErrHandler` 段开头先恢复两个开关，再提示与记日志 |
| `On Error Resume Next` 盖住整段代码，出错也不知道 | Resume Next 会吞掉所有错误，且事后没检查 `Err.Number` | 只在单个可预期操作上用，用完立即检查 `Err.Number` 并 `On Error GoTo 0` 恢复 |
| 日志文件写不进去 / 找不到 ErrorLog.txt | `ThisWorkbook.Path` 为空（新工作簿未保存）或目录无写权限 | 先保存工作簿再运行；或把日志路径改到固定目录（如 `Environ("TEMP")`） |
| 自定义错误号与系统错误冲突 | 自定义用了 512 以下的号，与系统错误号重叠 | 自定义错误号一律用 513 起（系统保留 0–512） |

## 7. 相关教程

- 同栏目上一篇：[04-examples.md](04-examples.md)（VBA 实战实例：8 个可直接运行的实例）
- 同栏目相关：[01-basics.md](01-basics.md)（VBA 基础：变量、循环、Sub/Function）、[02-objects.md](02-objects.md)（VBA 对象模型：Range、With、事件）、[03-keywords.md](03-keywords.md)（VBA 关键字大全：On Error、Err、Resume 速查）
- 跨栏目相关：[Power Query 教程](../powerquery/01-power-query.md)（数据清洗用 Power Query、自动化用 VBA，两者互补）

## 8. 练习题

1. `frmEntry.Show` 与 `frmEntry.Show vbModeless` 在行为上有什么区别？各适合什么场景？
2. 类模块 `clsEmployee` 里，`Salary` 写入 -100 时会发生什么？为什么校验要写在 `Property Let` 里而不是 `Property Get` 里？
3. 把工具打包成加载宏（.xlam）需要哪几步？为什么要保留一份 .xlsm 源文件？
4. 业务过程里 `On Error GoTo ErrHandler` 文末的 `ErrHandler:` 段为什么要先恢复 `Application.ScreenUpdating = True` 再提示错误？
5. 用本篇示例 2 的 3 名员工（月薪 8000、9500、12000）算：合计月薪、平均月薪（保留 2 位小数）、3 人年薪合计各是多少？
6. 某工具要处理 10 个文件，其中 2 个因路径错误触发 1004 被 `HandleError` 记入日志，其余 8 个成功。日志里应有几行记录？成功处理的文件数与日志行数之和是多少？

## 9. 参考答案

<details>
<summary>点击查看答案</summary>

1. `Show` 默认模态（vbModal=1）：用户必须关闭窗体才能操作工作表，之后的代码暂停执行，适合录入、确认类对话框；`Show vbModeless` 非模态：可一边开窗体一边操作表格，之后的代码立即继续执行，适合对照着表填写、常驻辅助面板。记忆点：模态"先关窗再干活"，非模态"窗表两不误"。

2. 触发 `Err.Raise 513`，错误描述"工资不能为负数"，写入被拒绝，`pSalary` 保持原值。校验必须写在 `Property Let`（写入入口）里：`Property Get` 只在读取时执行，写操作根本不经过它，校验写在那里等于没写。

3. 步骤：代码在普通工作簿里调通 → 调试 → 编译无错 → 文件 → 另存为 → 保存类型选 Excel 加载宏 (*.xlam) → 文件 → 选项 → 加载项 → 管理选 Excel 加载项 → 转到 → 勾选安装 → 重启 Excel 验证。保留 .xlsm 源文件的原因：.xlam 的工作表被隐藏、保存路径默认在加载宏目录，直接改 .xlam 容易丢改动；改源文件再重新另存覆盖才是可靠流程。

4. 因为业务过程开头关掉了 `ScreenUpdating`（批量操作提速），如果出错时直接退出而不恢复，Excel 会一直不刷新屏幕，看起来像"假死"，后续所有操作都受影响。固定写法：`ErrHandler` 段第一件事就是把关掉的开关（`ScreenUpdating`、`EnableEvents`、`DisplayAlerts`）全部恢复，再提示与记日志。

5. 合计月薪 29500（8000+9500+12000）；平均月薪 9833.33（29500÷3=9833.333…，四舍五入到 2 位小数）；3 人年薪合计 354000（96000+114000+144000，或 29500×12）。验算：8000×12=96000，9500×12=114000，12000×12=144000，三数相加 354000。

6. 日志里应有 2 行记录（每个出错文件记一行，成功的不记日志）。成功处理文件数 8 与日志行数 2 之和是 10，正好等于待处理文件总数，可用这个等式自查"有没有文件被悄悄漏掉"。

</details>

## 10. 可下载示例文件

- [VBA-UserForm-Demo.bas](templates/VBA-UserForm-Demo.bas)（报销录入窗体 frmEntry 的窗体代码与显示窗体的标准模块过程；窗体代码需粘贴进窗体代码窗口，另见第 4.1 节）
- [VBA-EmployeeClass.bas](templates/VBA-EmployeeClass.bas)（员工类 clsEmployee 的类模块代码；导入时在 VBE 中存为类模块，另见第 4.2 节）
- [VBA-AddIn-Tools.bas](templates/VBA-AddIn-Tools.bas)（加载宏工具库 modAddIn：Auto_Open 与跨工作簿可用的工具过程，另见第 4.3 节）
- [VBA-ErrorHandler.bas](templates/VBA-ErrorHandler.bas)（集中错误处理模块 modErrorHandler：LogError、HandleError 与业务过程固定写法，另见第 4.4 节）
- 04 篇的 8 个实例另见 [04-examples.md](04-examples.md)的"可下载示例文件"一节（8 个 .bas 文件）。

> **说明**：.bas 文件是纯文本 VBA 模块，可用 VBE「文件 → 导入文件」导入，或用记事本打开复制；类模块文件内容需粘贴进类模块（.cls），不可直接作为标准模块导入。所有代码均未在 Excel 内执行验证，导入后请先在副本工作簿上试跑。
