' VBA-AddIn-Tools.bas
' 对应教程：docs/vba/05-advanced.md 第 4.3 节：加载宏（.xlam）
' 使用方法：在 VBE 中导入本文件，或新建标准模块后全文粘贴；调通后按教程第 2.3 节「文件 → 另存为 → Excel 加载宏 (*.xlam)」保存，再在「文件 → 选项 → 加载项」里勾选安装。
' 说明：本文件代码经语法扫描核对，未在 Excel 内执行验证（无 Excel 环境）；导入后请先在副本工作簿上试跑。
Attribute VB_Name = "modAddIn"
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
