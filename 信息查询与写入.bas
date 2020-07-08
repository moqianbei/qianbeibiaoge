Sub 读取身份证信息()

MsgBox ("查询到的信息如下：" & Chr(10) _
& Chr(10) & "籍贯： " & JSHENFENZHENG(Selection, 1) _
& Chr(10) & "生日： " & Format(Chr(9) & JSHENFENZHENG(Selection, 2), "YYYY-MM-DD;@") _
& Chr(10) & "年龄： " & JSHENFENZHENG(Selection, 3) _
& Chr(10) & "属相： " & JSHENFENZHENG(Selection, 4) _
& Chr(10) & "星座： " & JSHENFENZHENG(Selection, 5) _
& Chr(10) & "性别： " & JSHENFENZHENG(Selection, 6) _
& Chr(10) & "校验： " & JSHENFENZHENG(Selection, 7) _
& Chr(10))

End Sub

Sub 生成中国各区域名称表()

msg = MsgBox("此操作可能会持续10~20秒，请确认", vbOKCancel)
If msg = 1 Then
    
    actsht = ActiveSheet.Name
    t = Timer
    Application.DisplayAlerts = False
    For Each sht In Sheets
        If sht.Name = "中国二级区划名称表（勿删）" Or sht.Name = "中国三级区划名称表（勿删）" Then sht.Delete
    Next
    Application.DisplayAlerts = True
    
    Call 删除出错名称
    
    If msg <> 1 Then
        Exit Sub
    Else
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Dim str
        '从表格中获取数据
        arr = QUYUDDMA2.Range("A2:E6717")
        arow = UBound(arr) '行数
        
        Dim dic1, dic2, dic3, dic4, dic5
        'Set dic1 = CreateObject("Scripting.Dictionary")
        'Set dic2 = CreateObject("Scripting.Dictionary")
        Set dic3 = CreateObject("Scripting.Dictionary")
        Set dic4 = CreateObject("Scripting.Dictionary")
        Set dic5 = CreateObject("Scripting.Dictionary")
        
        For i = LBound(arr) To UBound(arr)
        
        '    dic1(ARR(i, 1)) = 1 '代码列
        '    dic2(ARR(i, 2)) = 2 '完整区域列
            dic3(arr(i, 3)) = 3 '一级列
            dic4(arr(i, 3) & "_" & arr(i, 4)) = 4 '直接将省市一级连接，不添加其他符号，方便后续名称命名
            dic5(arr(i, 5)) = 5
            
        Next
        
        dic3cou = dic3.Count
        dic4cou = dic4.Count
        
        Set sht1 = Sheets.Add(before:=ActiveWorkbook.Sheets(1)) '二级INDIRECT
        Set sht2 = Sheets.Add(before:=ActiveWorkbook.Sheets(1))
        sht1.Name = "中国二级区划名称表（勿删）"
        sht2.Name = "中国三级区划名称表（勿删）"
    
        sht1.Range("a1").Resize(1, dic3cou) = dic3.keys '二级列表表头：北京市、天津市
        sht2.Range("a1").Resize(1, dic4cou) = dic4.keys '北京市市辖区
    
        acol1 = sht1.Range("A1").End(xlToRight).Column
        acol2 = sht2.Range("A1").End(xlToRight).Column
        
        For i = 1 To acol1
            m = 2
            str1 = sht1.Cells(1, i)
            
            For j = 1 To arow
                If arr(j, 3) = str1 And arr(j, 4) <> "" Then
                    sht1.Cells(m, i) = arr(j, 4)
                    m = m + 1
                End If
            Next
        
            sht1.Select
            sht1.Range(Cells(2, i), Cells(2, i).End(xlDown)).Select
            Selection.RemoveDuplicates Columns:=1, HEADER:=xlNo '去重
            
            If sht1.Cells(2, i) <> "" Then  '防止出现xldown为很大数值
                sht1.Range(sht1.Cells(1, i), sht1.Cells(1, i).End(xlDown)).Select
                Selection.CreateNames Top:=True, Left:=False, Bottom:=False, Right:=False '添加到名称管理器
            End If
        
        Next
        
        For i = 1 To acol2
            N = 2
            str2 = sht2.Cells(1, i)
            For j = 1 To arow
                If arr(j, 3) & "_" & arr(j, 4) = str2 And arr(j, 5) <> "" Then
                    sht2.Cells(N, i) = arr(j, 5)
                    N = N + 1
                End If
            Next
        
            sht2.Select
            sht2.Range(sht2.Cells(2, i), sht2.Cells(2, i).End(xlDown)).Select
            Selection.RemoveDuplicates Columns:=1, HEADER:=xlNo '去重
            
            If sht2.Cells(2, i) <> "" Then
                sht2.Range(sht2.Cells(1, i), sht2.Cells(1, i).End(xlDown)).Select
                Selection.CreateNames Top:=True, Left:=False, Bottom:=False, Right:=False '添加到名称管理器
            End If
        
        Next
        Sheets(actsht).Select
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        MsgBox "共用时" & Timer - t & "秒!"
    End If
    
    '隐藏这两个sheet
    sht1.Visible = False
    sht2.Visible = False
Else
    MsgBox "您选择了取消"
End If
End Sub

Sub 设置所选区域为一级代码()

arr = QUYUDDMA2.Range("A2:E6717")
Set dic3 = CreateObject("Scripting.Dictionary")
For i = LBound(arr) To UBound(arr)
        dic3(arr(i, 3)) = 3 '一级列
    Next

St = Join(dic3.keys, ",")

rng1 = Selection.Cells(1).Address(0, 0)
arow = Selection.Rows.Count

With Range(rng1).Resize(arow, 1).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
    xlBetween, Formula1:=St
    .ignoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .IMEMode = xlIMEModeNoControl
    .ShowInput = True
    .ShowError = True
End With

Range(rng1).Offset(0, 1).Resize(1, 2).NumberFormatLocal = "@"

Range(rng1).Offset(0, 1) = "=INDIRECT(" & rng1 & ")"
Range(rng1).Offset(0, 2) = "=INDIRECT(" & rng1 & "&""_""&" & Range(rng1).Offset(0, 1).Address(0, 0) & ")"

MsgBox "如果想设置二级、三级区域数据验证，需先生成各区域名称，然后要分别选中区域（单列），复制最上方单元格的值为数据验证的公式即可"

End Sub


