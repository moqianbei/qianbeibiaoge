
Sub 原位粘贴为值和源格式()

'只需定位到要执行的区域即可

    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
End Sub


Sub 原位粘贴为显示的值()

'只需定位到要执行的区域即可
  msg = MsgBox("此操作可能会耗费大量资源，且无法中断，请确认", vbOKCancel)
If msg = 1 Then
    Dim rng As Range
        For Each rng In Selection
            rng = rng.text
        Next
    Else
        Exit Sub
End If
End Sub

Sub 定位到选区的空值单元格()

    Selection.SpecialCells(xlCellTypeBlanks).Select

End Sub

Sub 数字转日期()

    Selection.NumberFormatLocal = "yyyy-mm-dd;@"

End Sub

Sub 添加前缀()

Dim rng As Range
On Error Resume Next
qmvv = InputBox("请输入要添加的前缀：")

For Each rng In Selection
    rng = qmvv & rng.value
Next

End Sub

Sub 添加后缀()

Dim rng As Range
On Error Resume Next
hzvv = InputBox("请输入要添加的前缀：")

For Each rng In Selection
    rng = rng.value & hzvv
Next

End Sub

Sub 文本型日期转真正日期()

Dim rng1 As Range
Set rng1 = Selection(1)

    Application.CutCopyMode = False
    Selection.TextToColumns destination:=rng1, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, other:=False, _
        FieldInfo:=Array(1, 5), TrailingMinusNumbers:=True

End Sub

Sub 去除选择区域重复值()

    Selection.RemoveDuplicates Columns:=1, HEADER:=xlNo

End Sub

Sub 清除单元格数据验证()

    Selection.Validation.Delete

End Sub
Sub 复制行高列宽格式()
'可跨sheet运行

Dim rng1 As Range, rng2 As Range
Dim i As Integer, j As Integer

On Error GoTo 100
Set rng1 = Application.InputBox("请选择要复制行高列宽的单元格", Type:=8)
On Error GoTo 100
Set rng2 = Application.InputBox(" 该格式应用在哪个区域？", Type:=8)
'flag标记是2，多次复制，1为依次复制，0位出错，不运行
If rng1.Cells.Count = 1 Then    '1→N 多次复制
    For Each rng In rng2
        rng.ColumnWidth = rng1.ColumnWidth
        rng.RowHeight = rng1.RowHeight
    Next
ElseIf rng2.Cells.Count = 1 Then    'N→1 依次复制
    For i = 1 To rng1.Rows.Count
        For j = 1 To rng1.Columns.Count
            rng2.Offset(i - 1, j - 1).ColumnWidth = rng1.Cells(i, j).ColumnWidth
            rng2.Offset(i - 1, j - 1).RowHeight = rng1.Cells(i, j).RowHeight
        Next
    Next
Else
    GoTo 100
End If
Exit Sub '提前结束，避免运行错误提示

100:
    MsgBox ("已取消选择/输入，程序已结束"):
    Exit Sub

End Sub
Sub 多列互转()

Dim arr

arr = Selection.Cells
Cot = Selection.Cells.Count
On Error Resume Next
Set rng = Application.InputBox("转置后的单元格放在区域的第一个单元格", Type:=8)
If Err Then
    MsgBox ("您未选择单元格，程序已结束")
    Err.Clear
    Exit Sub
Else

    On Error GoTo 100
    liehang = InputBox("请输入需要转换的列数或行数，以"",""分隔（lie,[hang]）")

    lie = Val(Split(liehang, ",")(0))
    If InStr(liehang, ",") <> 0 Then hang = Val(Split(liehang, ",")(1))

    If hang = "" Or hang = 0 Then hang = IIf(Cot / lie = Int(Cot / lie), Cot / lie, Int(Cot / lie + 1))
    If lie = "" Or lie = 0 Then lie = IIf(Cot / hang = Int(Cot / hang), Cot / hang, Int(Cot / hang + 1))

    On Error GoTo 100
    zixing = InputBox("请输入需要转换的形式，N先列后行，Z先行后列")

    m = 0
    N = 0
    Application.ScreenUpdating = False
    For i = LBound(arr) To UBound(arr)
        For j = LBound(arr, 2) To UBound(arr, 2)
            rng.Offset(m, N) = arr(i, j)
            
            If zixing = "n" Or zixing = "N" Then
                If m < hang - 1 Then
                    m = m + 1
                Else
                    N = N + 1
                    m = 0
                End If

            ElseIf zixing = "z" Or zixing = "Z" Then
                If N < lie - 1 Then
                    N = N + 1
                Else
                    m = m + 1
                    N = 0
                End If
            End If
        Next
    Next
End If
Application.ScreenUpdating = False
Exit Sub

100:
    MsgBox ("已取消选择/输入，程序已结束"):
    Exit Sub
End Sub
