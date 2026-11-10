' VBA-CustomFunctions.bas
' 对应教程：docs/vba/04-examples.md 实例 8：自定义函数（UDF）
' 使用方法：在 VBE 中导入本文件，或新建标准模块后全文粘贴；然后在单元格公式中使用，如 =GetBirth(A1)、=IsWeekend(A1)、=RMBUpper(A1)。
' 说明：本文件代码经语法扫描核对，未在 Excel 内执行验证（无 Excel 环境）；导入后请先在副本工作簿上试跑。
Attribute VB_Name = "modCustomFunctions"
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

' 让函数每次重算都执行
Function MyUDF() As String
    Application.Volatile
    MyUDF = "..."
End Function
