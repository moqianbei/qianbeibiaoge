Sub 选择全部sheet()
'Worksheets.Select
''不可取消sheets(1)的选中状态

Dim arr() '存储已显示的工作表名
'不选择第一个sheet
Dim sht As Worksheet
k = 0
For Each sht In Worksheets
    If sht.Visible = True Then
       k = k + 1
       ReDim Preserve arr(1 To k)
       arr(k) = sht.Name
    End If
Next

If UBound(arr) > 1 Then
    Sheets(arr(2)).Select '先选中可见的第二个sheet，也就取消了第一个sheet的选中状态
    For Each sht In Worksheets
        j = 0
        For i = LBound(arr) + 1 To UBound(arr) '排除第一张可见sheet
            If sht.Name = arr(i) Then
                j = 1
                Exit For
            End If
        Next
        If j = 1 Then sht.Select Replace:=False
    Next
Else
    MsgBox "仅有一张工作表可见"
End If

End Sub

Sub 删除未选中工作表()
Dim sht As Worksheet
Dim arr() As String
Dim i, cou As Integer, flag As Integer

cou = ActiveWindow.SelectedSheets.Count
i = 1

Dim msg As Byte
msg = MsgBox("此操作会永久删除所有其他已显示的sheet，请确定", vbOKCancel)

If msg <> 1 Then
    Exit Sub
Else

    Excel.Application.DisplayAlerts = False

    ReDim arr(1 To cou)
    For Each sht In ActiveWindow.SelectedSheets
        arr(i) = sht.Name
        i = i + 1
    Next
    
    For Each sht In Sheets
        If sht.Visible = xlSheetVisible Then '如果为显示状态
            flag = 0 '重置为0
            For i = LBound(arr) To UBound(arr)
                If sht.Name = arr(i) Then
                    flag = 1 '代表存在，不能删除
                    Exit For
                End If
            Next
            If flag = 0 Then
                On Error Resume Next
                sht.Delete
            End If
        End If
    Next
End If

Excel.Application.DisplayAlerts = True

End Sub

Sub 取消全部sheet隐藏()
    Dim sht As Worksheet
    For Each sht In Worksheets
        sht.Visible = xlSheetVisible
    Next
End Sub

Sub 删除出错名称()

Dim N As Name
For Each N In ActiveWorkbook.Names
    If InStr(N, "#REF!") Then N.Delete
Next

End Sub

Sub 根据选区建立工作表()

Dim rng As Range
Dim cou, flag As Byte
Dim i As Integer
Dim actsht, actsel
actsht = ActiveSheet.Name '保存当前活动sheet
actsel = Selection.Cells.Count

For Each rng In Selection
    flag = 0
    For i = 1 To Sheets.Count
        If Sheets(i).Name = (rng.value & "") Or rng.value & "" = "" Then
            flag = 1
            Exit For
        End If
    Next
    If flag = 0 Then
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = rng.value
    Else
        cou = cou + 1 '创建失败数目
    End If

Next
If cou <> 0 Then MsgBox ("有 " & cou & "/" & actsel & " Sheet因名称重复/空值创建失败")
Sheets(actsht).Select

End Sub

Sub 按列拆分表格()
Dim msg As Byte
Dim rng As Range
Set rng = ActiveCell

'选择的单元格不能为第1行
msg = MsgBox("请确定以下内容：" & Chr(10) _
& "1. 该Sheet为清单式表格(首行为标题行，首列不为空)" & Chr(10) _
& "2. 已选择拆分的依据列中的某单元格", vbOKCancel)

Set actsht = ActiveSheet '存储当前活动的sheet名称

If msg <> 1 Then
    Exit Sub
Else
    
    Dim i, col As Integer

    Excel.Application.DisplayAlerts = False
    With actsht
    
        '复制并删除重复值
        colnum = rng.Column '选择单元格的列号
        .Cells(2, colnum).Resize(.Cells(2, colnum).End(xlDown).Row + 1, 1).Copy Sheets.Add(before:=Sheets(1)).Range("a1") '复制到最左侧，此时该表为sheets(1)
        Sheets(1).Range("A1:A" & Sheets(1).[a1].End(xlDown).Row).RemoveDuplicates Columns:=1, HEADER:=xlNo  '去除重复值

        rnum = Sheets(1).[a1].End(xlDown).Row   '存储不重复值的个数
        '按名称新建表格于最后
        On Error GoTo 100
        For i = 1 To rnum
            Sheets.Add(After:=Sheets(Sheets.Count)).Name = Sheets(1).Range("a" & i)
        Next
        Sheets(1).Delete

        '复制对应内容到指定表格
        For i = Sheets.Count To Sheets.Count - rnum + 1 Step -1
            .Range(.[a1], .[a1].End(xlDown).End(xlToRight)).AutoFilter Field:=colnum, CRITERIA1:=Sheets(i).Name    '启动筛选
            .Range(.[a1], .[a1].End(xlDown).End(xlToRight)).Copy Sheets(i).[a1]
        Next

        '取消筛选
        .Select
        .Range(.[a1], .[a1].End(xlDown).End(xlToRight)).AutoFilter Field:=1
        Selection.AutoFilter
    End With
    Excel.Application.DisplayAlerts = True

End If
Exit Sub

100:

MsgBox "已有相同名称的Sheet" & "，请删除后重试"

End Sub

Sub 多工作簿的表合并在当前文件中()

Dim str()
Dim i As Integer
Dim wb, wb1 As Workbook
Dim sht, sht1 As Worksheet

shtname = ActiveSheet.Name
Set wb1 = ActiveWorkbook
Set sht1 = ActiveSheet

On Error Resume Next '加上这句防止用户点击取消发生的错误
str = Application.GetOpenFilename("Excel数据文件,*.xls*", Title:="请选择要合并的文件", MultiSelect:=True)
Application.ScreenUpdating = False

    For i = LBound(str) To UBound(str)
        Set wb = Workbooks.Open(str(i))
        For Each sht In wb.Sheets
            sht.Copy After:=wb1.Sheets(wb1.Sheets.Count)
            wb1.Sheets(wb1.Sheets.Count).Name = Split(wb.Name, ".")(0) & "_" & sht.Name '文件名中请勿出现“.”否则会有误
        Next
        wb.Close
    Next
Application.ScreenUpdating = True
Sheets(shtname).Select
End Sub

'需要事先选中标题行的下一个最左的数据，类似冻结窗格
'需要删除多余sheet，避免区域无法自动识别
'需要新建表格并复制表头，且空出第一列以放置各sheet名称

Sub 多sheet合并()

Dim newsht As Worksheet
Dim arr() As String
Dim i, con As Integer

Dim msg As Byte
msg = MsgBox("请确定已选择要合并的sheet", vbOKCancel)

If msg <> 1 Then
    Exit Sub
Else
    Application.ScreenUpdating = False
    con = ActiveWindow.SelectedSheets.Count
    ReDim arr(1 To con)
    i = 1
    For Each sht In ActiveWindow.SelectedSheets
        arr(i) = sht.Name
        i = i + 1
    Next
    
    On Error Resume Next '防止点击取消发生错误
    Set rng = Application.InputBox(" 请选择非标题的数据区域的最左上单元格", Type:=8)
    If rng = False Then
        MsgBox ("您未选择单元格，程序已结束")
        Exit Sub
    Else
        Excel.Application.DisplayAlerts = False
        
        arow = rng.Row
        acol = rng.Column
    
        Set newsht = Sheets.Add(before:=Sheets(1), Count:=1) '需要添加count，因为默认会添加你选择sheet的数量
    
        For i = LBound(arr) To UBound(arr)
            With Sheets(arr(i))
                .Range(.Cells(arow, acol), .Cells(arow, acol).End(xlToRight).End(xlDown)).Copy
                newsht.Select

                If i = LBound(arr) Then
                    Range("B65536").End(xlUp).Offset(arow - 1, 0).Select
                    ActiveSheet.Paste
                    newsht.Range(newsht.[A65536].End(xlUp).Offset(arow - 1, 0), newsht.[b65536].End(xlUp).Offset(0, -1)) = .Name
                Else
                    Range("B65536").End(xlUp).Offset(1, 0).Select
                    ActiveSheet.Paste
                    newsht.Range(newsht.[A65536].End(xlUp).Offset(1, 0), newsht.[b65536].End(xlUp).Offset(0, -1)) = .Name
                End If
            End With
        Next
        newsht.Select
        [a1] = "工作簿_工作表"
        [b1].Select '方便之后复制粘贴
        Sheets(arr(1)).Select
        Cells(1, acol).Select
        MsgBox "为防止自动复制遇到合并单元格出错，下面请您手动复制标题行"
        Excel.Application.DisplayAlerts = True
    End If
    Application.ScreenUpdating = True
End If

End Sub

Sub 活动表另存为工作簿文件()

Dim sht As Worksheet
On Error Resume Next
filepath = strFolder("请选择要保存到哪个文件夹")
If filepath = "" Then Exit Sub
Dim houzhui As String
Dim arr As Variant

'获取当前文件后缀名，防止文件名中有"."
arr = Split(ActiveWorkbook.Name, ".")
If arr(UBound(arr)) <> ActiveWorkbook.Name Then
    houzhui = arr(UBound(arr))
Else
    houzhui = "xlsx" '如果是新建的工作簿且未保存，设置默认后缀名
End If

'关闭屏幕更新，隐藏文件保存的过程
Application.ScreenUpdating = False

For Each sht In ActiveWindow.SelectedSheets

    sht.Copy '该方法会直接复制到新建的工作簿中，即新工作簿文件为之后激活的窗口
    ActiveWorkbook.SaveAs fileName:=filepath & sht.Name & "." & houzhui
    ActiveWorkbook.Close
Next

Application.ScreenUpdating = True

End Sub

Sub 拉sheet清单()

    On Error Resume Next '防止点击取消发生错误
    msg = MsgBox("请确认已选择需要生成目录的Sheet", vbOKCancel)
    If msg = 1 Then
        Dim shtarr()
        cou = ActiveWindow.SelectedSheets.Count
        ReDim shtarr(1 To cou)  '将所选sheet存入数组备用
        i = 1   '设置从1开始存储
        For Each sht In ActiveWindow.SelectedSheets
            Set shtarr(i) = sht
            i = i + 1
        Next
        
        On Error GoTo 100 '防止点击取消发生错误
        texttype = MsgBox("请选择要展示的文字，Yes为Sheet表名，No为单元格文字内容，Cancel为""Sheet1_B2""", vbYesNoCancel)

        Set rng = Application.InputBox("请选择你要链接到哪个单元格", Type:=8)
        If Err.Number = 13 Then GoTo 100    '如果点击了取消

        Dim textvalue   '显示的文字
        rngaddress = rng.Address(0, 0)  '获取相对地址如"B2"
        Set mulusht = Sheets.Add(before:=Sheets(1)) '   生成新sheet以存放目录
        With mulusht
            For i = LBound(shtarr) To UBound(shtarr)
                Select Case texttype
                Case 6  '   YES
                    textdis = shtarr(i).Name
                Case 7  'NO
                    textdis = shtarr(i).Range(rngaddress).value
                Case 2  'CANCEL
                    textdis = shtarr(i).Name & "_" & shtarr(i).Range(rngaddress).value
                End Select

                .Hyperlinks.Add Anchor:=Range("a" & i), Address:="", SubAddress:= _
                    shtarr(i).Name & "!" & rngaddress, ScreenTip:="莫浅北屏幕提示", TextToDisplay:=textdis

            Next
        End With
    Else
        GoTo 100
    End If
    Exit Sub

100:
    MsgBox ("已取消操作")
End Sub
