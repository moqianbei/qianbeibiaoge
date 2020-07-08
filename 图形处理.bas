
Sub 批量插入图片()

Dim rng1, rng2 As Range

On Error Resume Next '为防止取消选择后报错
Set rng1 = Selection.Cells

msg = MsgBox("请确认已选择图片路径（可先使用查找功能）", vbOKCancel)
If msg = vbCancel Then
    MsgBox "操作已取消"
    Exit Sub
Else

    For Each rng In rng1
        If rng = "" Then
            MsgBox ("选中区域有空单元格，程序已结束")
            Exit Sub
        End If
    Next

    '如果没有空单元格

    picdir = InputBox("请输入要偏移的位置，以英文逗号分隔(下移y行,右移x列)")
    x = Val(Split(picdir, ",")(0))
    y = Val(Split(picdir, ",")(1))

    For Each rng In rng1

        Set rng2 = rng.Offset(x, y)

        On Error Resume Next
        Dim shp As Shape
        Set shp = ActiveSheet.Shapes.AddPicture(rng.value, msoFalse, msoCTrue, rng2.MergeArea.Left, rng2.MergeArea.Top, rng2.MergeArea.width, rng2.MergeArea.height) '可匹配合并单元格
        shp.Placement = xlMoveAndSize '随单元格大小和位置改变
    Next
End If
End Sub
Sub 清除所有图形()

m = MsgBox("请确定清除该sheet中所有图形？该操作不可撤销！", vbOKCancel)

If m <> 1 Then
    Exit Sub
Else

    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        shp.Delete
    Next
End If
End Sub

Sub 选择含有图片路径的单元格()

    Dim firstAddress As String, c As Range, rALL As Range
    With ActiveSheet.usedRange
        Set c = .Find(":\", LookIn:=xlValues, LookAt:=xlPart)
        If Not c Is Nothing Then
            Set rALL = c
            firstAddress = c.Address
            Do
                Set rALL = Union(rALL, c)
                ActiveSheet.Range(c.Address).Activate
                Set c = .FindNext(c)

            Loop While Not c Is Nothing And c.Address <> firstAddress
        End If
        '.Activate   '未找到会全选已使用区域
        If Not rALL Is Nothing Then rALL.Select
    End With
End Sub
