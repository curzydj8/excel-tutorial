' VBA-EmployeeClass.bas
' 对应教程：docs/vba/05-advanced.md 第 4.2 节：类模块（员工类 clsEmployee）
' 使用方法：本文件分两部分：A 部分（类代码）需粘贴进类模块（VBE「插入 → 类模块」，名称改为 clsEmployee），不能作为标准模块导入；B 部分（TestEmployeeClass 测试过程）放在标准模块，按 F5 运行，立即窗口核对合计 29500、平均 9833.33。
' 说明：本文件代码经语法扫描核对，未在 Excel 内执行验证（无 Excel 环境）；导入后请先在副本工作簿上试跑。
' ===== A 部分：类模块 clsEmployee 的代码（粘贴进类模块）=====
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

' ===== B 部分：标准模块测试代码 =====
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
