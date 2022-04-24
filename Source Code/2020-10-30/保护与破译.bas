Dim autosave_workname, autosave_houzhui, autosave_filename, autosave_filepath

Sub 旧Excel文件破解()

Dim fileName
fileName = Application.GetOpenFilename("Excel文件,*.xls;*.xla;*.xlt", , "旧格式文件破解")
If Dir(fileName) = "" Then
    MsgBox "没找到相关文件,请重新设置"
    Exit Sub
Else
FileCopy fileName, fileName & ".bak" '备份文件
End If

Dim GetData As String * 5
Open fileName For Binary As #1
Dim CMGs As Long
Dim DPBo As Long
For i = 1 To LOF(1)
    Get #1, i, GetData
    If GetData = "CMG=""" Then CMGs = i
    If GetData = "[Host" Then DPBo = i - 2: Exit For
Next
If CMGs = 0 Then
    MsgBox "请先对VBA编码设置一个保护密码...", 32, "提示"
Exit Sub
End If

Dim St As String * 2
Dim s20 As String * 1
'取得一个0D0A十六进制字串
Get #1, CMGs - 2, St
'取得一个20十六制字串
Get #1, DPBo + 16, s20
'替换加密部份机码
For i = CMGs To DPBo Step 2
    Put #1, i, St
Next
'加入不配对符号
If (DPBo - CMGs) Mod 2 <> 0 Then
    Put #1, DPBo + 1, s20
End If
MsgBox "文件解密成功......", 32, "提示"
Close #1
End Sub

Sub xlam文件转xls()
    Dim strFile, wb As Workbook
    strFile = Application.GetOpenFilename(FileFilter:="Micrsoft Excel文件(*.xlam), *.xlam")
    If strFile = False Then Exit Sub
    With Workbooks.Open(strFile)
        .IsAddin = False
        .SaveAs fileName:=Replace(strFile, "xlam", "xls"), FileFormat:=xlExcel8
        .Close
    End With
End Sub


Sub CSV批量转xlsx()
'csv文件不能超过1048576行，否则会出错

Application.ScreenUpdating = False

Dim str()
Dim i As Integer
Dim wb As Workbook

On Error Resume Next '加上这句防止用户点击取消发生的错误
str = Application.GetOpenFilename("csv文件(*.csv),*.csv", Title:="请选择要转换的文件", MultiSelect:=True)

For i = LBound(str) To UBound(str)
    Set wb = Workbooks.Open(str(i), ReadOnly:=True)
    '保存为默认工作簿+常规工作簿文件
    wb.SaveAs Replace(str(i), ".csv", ""), IIf(Application.VERSION >= 12, xlWorkbookDefault, xlWorkbookNormal)
    wb.Close
Next

Application.ScreenUpdating = True
End Sub

Sub 设置可编辑区域()
Dim str As String
On Error GoTo 100
Selection.Locked = False
Selection.FormulaHidden = False
str = InputBox("请输入加密密码（可为空）")
MsgBox ("请记住你的工作表加密密码：" & str)
ActiveSheet.Protect Password:=str, DrawingObjects:=True, Contents:=True, Scenarios:=True
ActiveSheet.EnableSelection = xlUnlockedCells
Exit Sub
100:
MsgBox ("请先设置单元格锁定/撤销工作表保护！")
End Sub

Sub 设置禁止编辑区域()

End Sub
Sub 撤销工作表保护()

On Error GoTo 100
ActiveSheet.Unprotect
ActiveSheet.Cells.Locked = True '所有单元格锁定状态恢复默认
MsgBox ("已撤销工作表密码保护！")
Exit Sub
100:
MsgBox ("您输入的密码不正确！")

End Sub

Sub 保护工作簿结构()

Dim str As String

On Error GoTo 100
str = InputBox("请输入工作簿保护密码（可为空）")
ActiveWorkbook.Protect Password:=str, Structure:=True, Windows:=False
MsgBox ("请记住你的工作簿保护密码：" & str)
Exit Sub
100:
MsgBox ("请先解除工作簿保护！")
End Sub

Sub 撤销工作簿结构保护()

Dim str As String
str = InputBox("请输入解锁工作簿结构保护密码")
On Error GoTo 100
ActiveWorkbook.Unprotect str
MsgBox ("已解除工作簿结构保护！")
Exit Sub
100:
MsgBox ("您输入的密码不正确！")
End Sub
