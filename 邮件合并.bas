
Sub 开始邮件合并()

msg = MsgBox("请确保以下内容：" & Chr(10) & _
"使用>常规<表示常规型数据，" & Chr(10) & _
"使用》文本《表示文本型数据，" & _
"使用》图片<表示图片数据" & Chr(10) & _
"可借助添加前/后缀功能使用" _
& Chr(10) & Chr(10) & "将模板表和来源表命名为""模板表""、""来源表""" _
& Chr(10) & "模板表内容不超过A1:AZ100" _
& Chr(10) & "来源表内容从A1开始，且第1行为标题行，图片路径为完整路径（D:\IMG\name.jpg）", vbOKCancel)

If msg <> 1 Then
    Exit Sub
Else
    Sheets("来源表").Select
    On Error Resume Next '加上这句防止用户点击取消发生的错误
    shtname = InputBox("请问你要按照来源表的哪一列命名？（数字）")
    
    t = Timer
    Application.ScreenUpdating = False

    If shtname <> "" Then

        Dim arr(), crr(1 To 100 * 52)
        Dim shtRow, shtCol
        Dim i, j, k, g, h

        shtRow = Sheets("来源表").[a1].End(xlDown).Row
        shtCol = Sheets("来源表").[a1].End(xlToRight).Column

        arr = Sheets("来源表").[a1].Resize(shtRow, shtCol).value
        k = 0
        For i = 1 To 100
            For j = 1 To 52
                If Sheets("模板表").Cells(i, j) <> "" Then
                    If Left(Sheets("模板表").Cells(i, j), 1) = ">" And Right(Sheets("模板表").Cells(i, j), 1) = "<" Then
                        g1 = Sheets("模板表").Cells(i, j).value
                        For h1 = LBound(arr, 2) To UBound(arr, 2)
                            If arr(1, h1) = Mid(g1, 2, Len(g1) - 2) Then
                                g1 = h1
                                Exit For
                            End If
                        Next
                        k = k + 1
                        crr(k) = i & "," & j & "," & g1 & ",G/通用格式"
                        
                    ElseIf Left(Sheets("模板表").Cells(i, j), 1) = "》" And Right(Sheets("模板表").Cells(i, j), 1) = "《" Then
                        g2 = Sheets("模板表").Cells(i, j).value
                        For h2 = LBound(arr, 2) To UBound(arr, 2)
                            If arr(1, h2) = Mid(g2, 2, Len(g2) - 2) Then
                                g2 = h2
                                Exit For
                            End If
                        Next
                        k = k + 1
                        crr(k) = i & "," & j & "," & g2 & ",@"

                    ElseIf Left(Sheets("模板表").Cells(i, j), 1) = "》" And Right(Sheets("模板表").Cells(i, j), 1) = "<" Then
                        g3 = Sheets("模板表").Cells(i, j).value
                        For H3 = LBound(arr, 2) To UBound(arr, 2)
                            If arr(1, H3) = Mid(g3, 2, Len(g3) - 2) Then
                                g3 = H3
                                Exit For
                            End If
                        Next
                        k = k + 1
                        crr(k) = i & "," & j & "," & g3 & ",;;;"
                    End If
                End If
            Next
        Next


        '开始复制
        For h = 2 To shtRow
            Dim sht As Worksheet
            Sheets("模板表").Copy After:=Sheets(Sheets.Count)
            Set sht = ActiveSheet
            For g = 1 To k
                m = Val(Split(crr(g), ",")(0))
                N = Val(Split(crr(g), ",")(1))
                p = Val(Split(crr(g), ",")(2)) '来源表的第几列，可能为数值和文件路径
                Q = Split(crr(g), ",")(3) '单元格格式代码

                With sht.Cells(m, N)
                    If Q <> ";;;" Then
                        .MergeArea.NumberFormatLocal = Q '设置单元格格式
                        .MergeArea.value = arr(h, p) '设置单元格数值
                    Else
                        On Error Resume Next
                        Dim shp, shp1 As Shape

                        For Each shp1 In sht.Shapes
                            shp1.Delete
                        Next

                        Set shp = sht.Shapes.AddPicture(arr(h, p), False, True, .Left, .Top, .MergeArea.width, .MergeArea.height)
                        shp.Placement = xlMoveAndSize '随单元格大小和位置改变
                    End If
                End With
            Next
            sht.Name = arr(h, shtname) '命名列不能重复
        Next
        Application.ScreenUpdating = False
        Sheets("来源表").Select
        MsgBox "完成" & shtRow - 1 & "张表，共用时" & Timer - t & "秒！"
    End If
End If
End Sub

Sub 设置模板单元格为文本()

For Each rng In Selection
    If rng <> "" Then rng.value = "》" & rng.value & "《"
Next

End Sub

Sub 设置模板单元格为常规()

For Each rng In Selection
    If rng <> "" Then rng.value = ">" & rng.value & "<"
Next

End Sub
Sub 设置模板单元格为图片()

For Each rng In Selection
    If rng <> "" Then rng.value = "》" & rng.value & "<"
Next

End Sub

Sub 更改表名为来源表()

ActiveSheet.Name = "来源表"

End Sub

Sub 更改表名为模板表()

ActiveSheet.Name = "模板表"

End Sub

Sub 多表复制到一张表()

Dim newsht As Worksheet
Dim arr()
Dim rng As Range
Dim msg As Byte
Dim i, j, k, con As Integer

msg = MsgBox("请确保已选中需要合并的表格", vbOKCancel)

If msg <> 1 Then
    Exit Sub
Else

    Set rng = Application.InputBox(" 复制各表格的哪个区域？", Type:=8)
    Dim rngaddress
    Dim colcount, rowcount, hangshu As Integer
    If VBA.Strings.InStr("!", rng.Address) <> 0 Then
        rngaddress = Split(rng.Address, "!")(1)
    Else
        rngaddress = rng.Address
    End If
    
    colcount = rng.Columns.Count
    rowcount = rng.Rows.Count
    hangshu = InputBox("请输入每行放多少个表格")
    
    'Excel.Application.DisplayAlerts = False

    con = ActiveWindow.SelectedSheets.Count
    ReDim arr(1 To con)

    i = 1
    For Each sht In ActiveWindow.SelectedSheets
        arr(i) = sht.Name
        i = i + 1
    Next

    Set newsht = Sheets.Add(before:=Sheets(1), Count:=1)
    shtname = newsht.Name

    i = 1
    j = 1
    For k = LBound(arr) To UBound(arr)
        Sheets(arr(k)).Range(rngaddress).Copy Sheets(shtname).Cells(i, j)

        If j <= (hangshu - 1) * colcount Then
            j = j + colcount
        Else
            i = i + rowcount
            j = 1
        End If
    Next

End If

Excel.Application.DisplayAlerts = True

End Sub

