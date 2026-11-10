' VBA-SendMailViaOutlook.bas
' 对应教程：docs/vba/04-examples.md 实例 7：发送邮件（Outlook）
' 使用方法：在 VBE 中导入本文件，或新建标准模块后全文粘贴；先建好“收件人”表（A列邮箱、B列姓名），测试时先把 .Send 改成 .Display 预览，确认无误再改回 .Send；需本机安装并配置 Outlook。
' 说明：本文件代码经语法扫描核对，未在 Excel 内执行验证（无 Excel 环境）；导入后请先在副本工作簿上试跑。
Attribute VB_Name = "modSendMail"
Option Explicit

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
