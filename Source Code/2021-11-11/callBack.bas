Attribute VB_Name = "callBack"
Option Explicit
'=======================
'
'     公共内部函数
'
'=======================

Private Function hasWorkSheet(ByVal shtName As String) As Boolean
    Dim wsht
    On Error Resume Next
    Set wsht = Worksheets(shtName)
    If Err.Number = 0 Then
        hasWorkSheet = True
    Else
        hasWorkSheet = False
    End If
End Function



'=======================
'  单 元 格 数 字 格 式
'=======================

'单元格显示为万亿KM
'------- https://club.excelhome.net/thread-1312477-1-1.html -------------------------------
'1、万位保留1位小数：自定义格式应设置为0!.0,"万"或0"."#,"万" 或0"."0,"万"
'2、万位保留0位小数，但有千分符：自定义格式应设置为#,##0,,"万"%
'3、万位保留1位小数，且有千分符：自定义格式应设置为#,##0.0,,"万"%
'4、万位保留2位小数，且有千分符：自定义格式应设置为#,##0.00,,"万"%
'如果不需要单位，去掉“"万"”（单位及对应双引号）

'关于#,##0,,%自定义格式的设置说明：
'1.输入#,##0,,%→鼠标将光标移至%前→按住Alt键的同时，依次按选小键盘的数字1和0（或者使用Ctrl+J）
'2.将选定对象设定为“自动换行”隐藏最后的“%”
'--------------------------------------------------------------------------------------------
Private Sub digitalFormat(numberscale As String, Optional accuracy As Byte = 2, Optional hasSeparator As Boolean = False)
    Select Case numberscale
        Case "k", "K"
            Select Case accuracy
                Case 0
                    Selection.NumberFormatLocal = "0,""K"""     '"在VBA中需要使用""转义
                Case 1
                    Selection.NumberFormatLocal = "0.0,""K"""
                Case 2
                    Selection.NumberFormatLocal = "0.00,""K"""
                Case Else
            End Select
        Case "M", "m"
            Select Case accuracy
                Case 0
                   Selection.NumberFormatLocal = "0,,""M"""
                Case 1
                   Selection.NumberFormatLocal = "0.0,,""M"""
                Case 2
                   Selection.NumberFormatLocal = "0.00,,""M"""
                Case Else
            End Select
        Case "W", "w"
            If hasSeparator Then    '有千分符
                Select Case accuracy
                    Case 0
                        Selection.NumberFormatLocal = "#,##0,,""万""" & Chr(10) & "%"   'chr(10)在单元格内回车（Ctrl+J）
                    Case 1
                        Selection.NumberFormatLocal = "#,##0.0,,""万""" & Chr(10) & "%"
                    Case 2
                        Selection.NumberFormatLocal = "#,##0.00,,""万""" & Chr(10) & "%"
                    Case Else
                End Select

            Else                    '无千分符
                Select Case accuracy
                    Case 0
                        
                        Selection.NumberFormatLocal = "0,,""万""" & Chr(10) & "%"
                    Case 1
                        Selection.NumberFormatLocal = "0,,.0""万""" & Chr(10) & "%"
                    Case 2
                        Selection.NumberFormatLocal = "0,,.00""万""" & Chr(10) & "%"
                    Case Else
                End Select
            End If
        Case "HM", "hm"
            Selection.NumberFormatLocal = "0!.00,,""亿"""
    End Select

    Selection.WrapText = True   '自动换行
    Selection.EntireRow.AutoFit '自动行高

End Sub
Sub digitalK0(control As IRibbonControl)
    Call digitalFormat("K", 0)
End Sub
Sub digitalK1(control As IRibbonControl)
    Call digitalFormat("K", 1)
End Sub
Sub digitalK2(control As IRibbonControl)
    Call digitalFormat("K", 2)
End Sub
Sub digitalW0(control As IRibbonControl)
    Call digitalFormat("W", 0)
End Sub
Sub digitalW1(control As IRibbonControl)
    Call digitalFormat("W", 1)
End Sub
Sub digitalW2(control As IRibbonControl)
    Call digitalFormat("W", 2)
End Sub
Sub digitalM0(control As IRibbonControl)
    Call digitalFormat("M", 0)
End Sub
Sub digitalM1(control As IRibbonControl)
    Call digitalFormat("M", 1)
End Sub
Sub digitalM2(control As IRibbonControl)
    Call digitalFormat("M", 2)
End Sub
Sub digitalHM(control As IRibbonControl)
    Call digitalFormat("HM")
End Sub

'=======================
'  数 字 转 日 期 格 式
'=======================

Sub dateAcross(control As IRibbonControl)
    Selection.NumberFormatLocal = "yyyy-mm-dd;@"
End Sub
Sub dateText(control As IRibbonControl)
    Selection.NumberFormatLocal = "yyyy年mm月dd日;@"
End Sub

'=======================
'  单 元 格 文 本 处 理
'=======================
Sub addPrefix(control As IRibbonControl)
    Dim prefix As String
    On Error Resume Next
    prefix = InputBox("请输入要添加的前缀：")
    If Err Then Exit Sub
    Dim rng As Range
    For Each rng In Selection
        rng.Value = prefix & rng.Value
    Next
End Sub
Sub addSuffix(control As IRibbonControl)
    Dim suffix As String
    On Error Resume Next
    suffix = InputBox("请输入要添加的后缀：")
    If Err Then Exit Sub
    Dim rng As Range
    For Each rng In Selection
        rng.Value = rng.Value & suffix
    Next
    rng.EntireColumn.AutoFit
End Sub
'转为日期
Sub transferDate(control As IRibbonControl)

    Dim rng As Range
    Set rng = Selection
    Dim findArr() As Variant
    findArr = Array("年", "月", "~", "。")  '注意没有“日”
    Dim i As Long
    For i = 1 To rng.Columns.Count
        '使用分列方式转换字符串
        Selection.Columns(i).TextToColumns DataType:=xlDelimited, FieldInfo:=Array(1, 5), TrailingMinusNumbers:=True
              'xlDelimited使用分隔符分割
              'TextQualifier是使用单引号、双引号还是不使用引号作为文本限定符
        Selection.Columns(i).AutoFit    '自动调整列宽，防止出现####错误
    Next

    For i = LBound(findArr) To UBound(findArr)
        Selection.Replace What:=findArr(i), Replacement:="/", LookAt:=xlPart, _
            SearchOrder:=xlByRows, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Next
    Selection.Replace What:="日", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    Selection.NumberFormatLocal = "yyyy-mm-dd;@"

End Sub
'千万亿KM转真正数字()
Sub transferNumber(control As IRibbonControl)

    If MsgBox("请确认所选的每个单元格满足以下条件" & Chr(10) & _
      "    1.至多包含“万亿KM”中的一个单位，该单位位于末尾" & Chr(10) & _
      "    2.不能出现其他符号，如“,￥$”等", vbOKCancel) = vbOK Then

        Dim rng As Range
        For Each rng In Selection
            Dim str As String
            str = Right(rng.Value, 1)
    
            If UCase(str) = "K" Or str = "千" Then
                rng.Value = VBA.Strings.Left(rng, VBA.Strings.Len(rng) - 1) * 1000
                Selection.NumberFormatLocal = "0.00,""K"""
            ElseIf UCase(str) = "M" Then
                rng.Value = VBA.Strings.Left(rng, VBA.Strings.Len(rng) - 1) * 1000000
                rng.NumberFormatLocal = "0.00,,""M"""
            ElseIf str = "亿" Then
                rng.Value = VBA.Strings.Left(rng, VBA.Strings.Len(rng) - 1) * 100000000
                rng.NumberFormatLocal = "0!.00,,""亿"""
            ElseIf UCase(str) = "W" Or str = "万" Then
                rng.Value = VBA.Strings.Left(rng, VBA.Strings.Len(rng) - 1) * 10000
                rng.NumberFormatLocal = "0,,.00""万""" & Chr(10) & "%"
            Else
            End If

            Selection.WrapText = True
            Selection.EntireRow.AutoFit
        Next
    Else
    End If

End Sub

'=======================
'
'  单 元 格 数 据 验 证
'
'=======================

Sub identityCard(control As IRibbonControl)
    Dim rngNo1 As String
    rngNo1 = Selection.Cells(1).Address(0, 0)

    With Selection.Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=OR(AND(LEN(" & rngNo1 & ")=15,ISNUMBER(--VALUE(" & rngNo1 & "))),AND(LEN(" & rngNo1 & ")=18,ISNUMBER(--LEFT(" & rngNo1 & ",17)),OR(ISNUMBER(--RIGHT(" & rngNo1 & ",1)),RIGHT(" & rngNo1 & ",1)=""X"")))"
        .ignoreBlank = True
        .InputTitle = "身份证号"
        .InputMessage = "18位字符或15位字符"
        .ErrorTitle = "录入身份证出错"
        .ErrorMessage = "您输入的可能不符合身份证规范，请检查"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Sub phoneNumber(control As IRibbonControl)

    Dim rngNo1 As String
    rngNo1 = Selection.Cells(1).Address(0, 0)

    With Selection.Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=OR(AND(LEN(" & rngNo1 & ")=11,ISNUMBER(--VALUE(" & rngNo1 & "))),LEN(" & rngNo1 & ")=13)"
        .ignoreBlank = True
        .InputTitle = "手机号"
        .ErrorTitle = "11位字符可用两个空格或-分隔"
        .InputMessage = "手机号不正确"
        .ErrorMessage = "您输入手机号不符合规范，请检查"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Sub BankCard(control As IRibbonControl)
'-------- http://www.woshipm.com/pd/371041.html ------------------------

    Dim rngNo1 As String
    rngNo1 = Selection.Cells(1).Address(0, 0)

    With Selection.Validation
        .Delete
        .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=AND(LEN(" & rngNo1 & ")>13 , LEN(" & rngNo1 & ")<19,ISNUMBER(--VALUE(" & rngNo1 & ")))"
        .ignoreBlank = True
        .InputTitle = "银行卡号"
        .ErrorTitle = "13~19位数字字符"
        .InputMessage = "银行卡号不正确"
        .ErrorMessage = "您输入银行卡号不符合规范，请检查"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With

End Sub
Sub premierLevel(control As IRibbonControl)

    If Selection.Columns.Count > 1 Then
        MsgBox "请选择要设置一级区划的数据"
    Else
        '复制表格到当前工作簿
        On Error Resume Next
        Application.DisplayAlerts = False   '阻止是否替换当前表对话框
        If hasWorkSheet("中国所有区划名称表（勿删）") = False Then ThisWorkbook.premierSheet0.Copy before:=ActiveWorkbook.Worksheets(1)
        If hasWorkSheet("中国一级区划名称表（勿删）") = False Then ThisWorkbook.premierSheet1.Copy before:=ActiveWorkbook.Worksheets(1)
        If hasWorkSheet("中国二级区划名称表（勿删）") = False Then ThisWorkbook.premierSheet2.Copy before:=ActiveWorkbook.Worksheets(1)
        If hasWorkSheet("中国三级区划名称表（勿删）") = False Then ThisWorkbook.premierSheet3.Copy before:=ActiveWorkbook.Worksheets(1)
        Application.DisplayAlerts = True
        On Error GoTo 0

        '创建名称
        Application.DisplayAlerts = False   '阻止是否替换对话框
        ActiveWorkbook.Worksheets("中国二级区划名称表（勿删）").Cells.CreateNames Top:=True, Left:=False, Bottom:=False, Right:=False   '添加到名称管理器
        ActiveWorkbook.Worksheets("中国三级区划名称表（勿删）").Cells.CreateNames Top:=True, Left:=False, Bottom:=False, Right:=False
        Application.DisplayAlerts = True
    
        '隐藏sheet，不在前端显示
        ActiveWorkbook.Worksheets("中国一级区划名称表（勿删）").Visible = False
        ActiveWorkbook.Worksheets("中国二级区划名称表（勿删）").Visible = False
        ActiveWorkbook.Worksheets("中国三级区划名称表（勿删）").Visible = False
        
        '清除错误值
        Dim n As Name
        For Each n In ActiveWorkbook.Names
            If InStr(n, "#REF!") Then n.Delete
        Next
    
        Dim firstRng As String, sRow As Long
        firstRng = Selection.Cells(1).Address(0, 0)
        
        Dim pre1 As Variant, fm As String
        '返回r×1的二维数组
        pre1 = ActiveWorkbook.Worksheets("中国一级区划名称表（勿删）").Range(ActiveWorkbook.Worksheets("中国一级区划名称表（勿删）").[A1], ActiveWorkbook.Worksheets("中国一级区划名称表（勿删）").[A1].End(xlDown))
        Dim i As Long
        For i = LBound(pre1) To UBound(pre1) - 1
            fm = fm & pre1(i, 1) & ","
        Next
        fm = fm & pre1(UBound(pre1), 1)
    
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=fm
            .ignoreBlank = True '忽略空值
            .InCellDropdown = True  '提供下拉列表
            .InputTitle = "一级区划"
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True       '提示要输入什么信息
            .ShowError = False      '不提示错误信息
            .ignoreBlank = True
        End With
    
        Selection.Cells(1).Offset(0, 1).Resize(1, 2).NumberFormatLocal = "@"
        Selection.Cells(1).Offset(0, 1) = "=INDIRECT(" & firstRng & ")"
        Selection.Cells(1).Offset(0, 2) = "=INDIRECT(" & Selection.Cells(1).Address(0, 0) & "&""_""&" & Selection.Cells(1).Offset(0, 1).Address(0, 0) & ")"
    
        MsgBox "已完成一级区划设置！" & Chr(10) & "如果想设置二、三级区域数据，可分别选中，复制最上方单元格的值作为数据验证-序列的公式即可"

    End If

End Sub
'清除数据验证
Sub cleanVerification(control As IRibbonControl)
    Selection.Validation.Delete
End Sub

'========================
'
'  单 元 格 信 息 读 取
'
'========================
Sub readIdentityCard(control As IRibbonControl)
    If Selection.Count = 1 Then
        Dim v As String
        v = Selection.Value & ""
        Dim idmsg As New IDCard
        If idmsg.Info(v, 6) Like "*不*规范*" Then
            MsgBox idmsg.Info(v, 6)
        Else
            MsgBox "查询到的身份证信息如下：" & Chr(10) _
                & Chr(10) & "籍贯： " & idmsg.Info(v, 1) _
                & Chr(10) & "生日： " & Format(Chr(9) & idmsg.Info(v, 2), "YYYY-MM-DD;@") _
                & Chr(10) & "年龄： " & idmsg.Info(v, 3) _
                & Chr(10) & "属相： " & idmsg.Info(v, 4) _
                & Chr(10) & "星座： " & idmsg.Info(v, 5) _
                & Chr(10) & "性别： " & idmsg.Info(v, 6) _
                & Chr(10) & "校验： " & idmsg.Info(v, 7) _
                & Chr(10)
        End If
    Else
        MsgBox "只能选择一个单元格"
    End If
End Sub

Sub readBankCard(control As IRibbonControl)
' ------ 使用免费API: http://www.zhaotool.com/api/bank  ---------------
    If Selection.Count = 1 Then
        Dim v As String
        v = Selection.Value & ""
        Dim bankC As New BankCard
        MsgBox "查询到的银行卡信息如下：" & Chr(10) _
        & Chr(10) & "所在银行名称： " & bankC.Info(v, 1) _
        & Chr(10) & "银行卡的类型： " & bankC.Info(v, 2) _
        & Chr(10) & "所在银行电话： " & bankC.Info(v, 3) _
        & Chr(10) & "校验码正确性： " & bankC.Info(v, 4) _
        & Chr(10)
    Else
        MsgBox "只能选择一个单元格"
    End If
End Sub

Sub readPhoneNumber(control As IRibbonControl)

    If Selection.Count = 1 Then
        Dim p As New phoneCard
        MsgBox p.Info(Selection.Value & "")
    Else
        MsgBox "只能选择一个单元格"
    End If

End Sub
Sub readPostCode(control As IRibbonControl)

End Sub
'========================
'
'    原 位 性 粘 贴
'
'========================

'原位粘贴为值和源格式()
Sub pasteValue(control As IRibbonControl)
'只需定位到要执行的区域即可

    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False
    Application.CutCopyMode = False

End Sub
'原位粘贴为显示的值()
Sub pasteDisplayValue(control As IRibbonControl)
'只需定位到要执行的区域即可

    If MsgBox("此操作可能耗费一些资源，且无法中断，确认进行吗", vbOKCancel) = vbOK Then
        Dim rng As Range
        For Each rng In Selection
            rng.Value = rng.Text
        Next
    End If

End Sub
'粘贴行高列宽格式()
Sub pasteRowHColumnW(control As IRibbonControl)
'可跨sheet运行

    Dim rng1 As Range, rng2 As Range
    Dim i As Integer, j As Integer

    Set rng1 = Selection
    
    On Error Resume Next
    Set rng2 = Application.InputBox(" 该格式应用在哪个区域？", Type:=8)
    If Err.Number = 424 Then GoTo cancelSelect
    On Error GoTo 0
    
    If rng1.Cells.Count = 1 Then           '1→N 多次复制
        Dim rng As Range
        For Each rng In rng2
            rng.ColumnWidth = rng1.ColumnWidth
            rng.RowHeight = rng1.RowHeight
        Next
    ElseIf rng2.Cells.Count = 1 Then        'N→1 依次复制
        For i = 1 To rng1.Rows.Count
            For j = 1 To rng1.Columns.Count
                rng2.Offset(i - 1, j - 1).ColumnWidth = rng1.Cells(i, j).ColumnWidth
                rng2.Offset(i - 1, j - 1).RowHeight = rng1.Cells(i, j).RowHeight
            Next
        Next
    Else
        MsgBox "只能1→N或N→1，且N为连续区域"
    End If
    Exit Sub '提前结束，避免运行错误提示

cancelSelect:
    MsgBox ("已取消选择/输入，程序已结束"):
    Err.Clear
End Sub
'========================
'
'    图 片 图 形 处 理
'
'========================
Sub selectPicture(control As IRibbonControl)
'    Dim rng As Range, nowRng As Range, sRng() As Range, i As Long
'    Set nowRng = Selection
'    For Each rng In nowRng
'        If rng.value Like "*:\*jpg" Then
'            sRng(i) = rng
'            i = i + 1
'        End If
'    Next
'    For i = LBound(sRng) To UBound(sRng)
'
'    Next

End Sub
'批量插入图片()
Sub insertPicture(control As IRibbonControl)

    If MsgBox("请确认已设置完整路径（D:\folder\abc.jpg）", vbOKCancel) = vbOK Then

        Dim rng1 As Range
        Set rng1 = Selection.Cells '存储选区

        Dim picdir As String
        picdir = InputBox("请输入要偏移的位置，以英文逗号分隔(下移y行,右移x列)")
        If picdir <> "" Then
            Dim errRng As Range, errNumber As Long  '存储发生错位的单元格
            
            Dim x As Double, y As Double
            x = Val(Split(picdir, ",")(0))
            If InStr(picdir, ",") <> 0 Then y = Val(Split(picdir, ",")(1))
            Dim rng As Range, rng2 As Range
            Dim shp As Shape
            For Each rng In rng1
                Set rng2 = rng.Offset(x, y)
                If rng.Value <> "" Then
                    On Error GoTo cannotFind
                    Set shp = ActiveSheet.Shapes.AddPicture(rng.Value, msoFalse, msoCTrue, rng2.MergeArea.Left, rng2.MergeArea.Top, rng2.MergeArea.width, rng2.MergeArea.height) '可匹配合并单元格
                    shp.Placement = xlMoveAndSize '随单元格大小和位置改变
                End If
            Next
            If errNumber > 0 Then
                MsgBox "有" & errNumber & "个图片无法添加，已为您选中错误单元格"
                errRng.Select
            End If
        End If
        
    End If
    Exit Sub
cannotFind:
    If errNumber = 0 Then
        Set errRng = rng
    Else
        Set errRng = Union(errRng, rng) '加选区，方便之后选择后展示出来
    End If
    errNumber = errNumber + 1
    Resume Next
End Sub
'清除所有图形()
Sub deleteShape(control As IRibbonControl)
    Dim msg As VbMsgBoxResult
    msg = MsgBox("请确定清除该sheet中仅图片（Y）还是所有图形（N）？", vbYesNoCancel)
    If msg <> vbCancel Then
        Dim shp As Shape
        For Each shp In ActiveSheet.Shapes
            If msg = vbYes Then
                If shp.Type = msoPicture Then shp.Delete
            Else
                shp.Delete
            End If
        Next
    Else
    End If
End Sub
'定位到选区的空值单元格()
Sub locateNull(control As IRibbonControl)
    Selection.SpecialCells(xlCellTypeBlanks).Select
End Sub
'转换行列
Sub convertRowsColumns(control As IRibbonControl)
    
    Dim arr, cot As Long

    arr = Selection.Cells
    cot = Selection.Cells.Count

    Dim rng As Range
    Set rng = Application.InputBox("转置后的单元格放在区域的第一个单元格", Type:=8)
    If Err Then
        GoTo 100
    Else
        Dim liehang As String

        liehang = InputBox("请输入需要转换的列数或行数，以"",""分隔（lie,[hang]）")
        If Err Then GoTo 100
        
        Dim lie As Double, hang As Double

        lie = Val(Split(liehang, ",")(0))
        If InStr(liehang, ",") <> 0 Then hang = Val(Split(liehang, ",")(1))
        If hang = 0 Then hang = IIf(cot / lie = Int(cot / lie), cot / lie, Int(cot / lie + 1))
        If lie = 0 Then lie = IIf(cot / hang = Int(cot / hang), cot / hang, Int(cot / hang + 1))

        Dim zixing As String
        zixing = InputBox("请输入需要转换的形式，N先列后行，Z先行后列")
        If Err Then GoTo 100

        Dim n As Long, m As Long
        m = 0
        n = 0
        Application.ScreenUpdating = False
        Dim i As Long, j As Long
        For i = LBound(arr) To UBound(arr)
            For j = LBound(arr, 2) To UBound(arr, 2)
                rng.Offset(m, n) = arr(i, j)

                If zixing = "n" Or zixing = "N" Then
                    If m < hang - 1 Then
                        m = m + 1
                    Else
                        n = n + 1
                        m = 0
                    End If
    
                ElseIf zixing = "z" Or zixing = "Z" Then
                    If n < lie - 1 Then
                        n = n + 1
                    Else
                        m = m + 1
                        n = 0
                    End If
                End If
            Next
        Next
    End If
    Application.ScreenUpdating = False
    Exit Sub
    
100:
    MsgBox ("已取消选择/输入，程序已结束"):

End Sub
'去重
Sub removeDuplicates(control As IRibbonControl)
    Selection.removeDuplicates Columns:=1, Header:=xlNo
End Sub
'========================
'
'    工 作 表 处 理
'
'========================
'取消全部sheet隐藏()
Sub showAllWorksheets(control As IRibbonControl)
    Dim sht As Worksheet
    For Each sht In Worksheets
        sht.Visible = xlSheetVisible
    Next
End Sub

'选择除1外所有可见sheet()
Sub selectWorksheets(control As IRibbonControl)
    
    'Worksheets.Select：不可取消sheets(1)的选中状态
    
    Dim arr() '存储已显示的工作表名
    '不选择第一个sheet
    Dim k As Long
    k = 0
    Dim sht As Worksheet
    For Each sht In Worksheets
        If sht.Visible = True Then
           k = k + 1
           ReDim Preserve arr(1 To k) '重新定义数组大小，但保留之前的值
           arr(k) = sht.Name
        End If
    Next
    
    If UBound(arr) - LBound(arr) > 0 Then
        Sheets(arr(LBound(arr) + 1)).Select '先选中可见的第二个sheet，也就取消了第一个sheet的选中状态
        For Each sht In Worksheets
            Dim flag As Byte
            flag = False    '假设没有重复值
            Dim i As Long
            For i = LBound(arr) + 1 To UBound(arr) '排除第一张可见sheet
                If sht.Name = arr(i) Then
                    flag = True
                    Exit For
                End If
            Next
            If flag = True Then sht.Select Replace:=False   '加选，不替换之前的选择
        Next
    Else
        MsgBox "仅有一张工作表可见"
    End If

End Sub

Sub deleteNotSelectedWorksheets(control As IRibbonControl)
    
    If MsgBox("此操作会永久删除所有其他已显示的sheet，请确定", vbOKCancel) = vbOK Then
    
        Application.ScreenUpdating = False

        Dim arr() As String
        
        Dim cou As Long
        cou = ActiveWindow.selectedsheets.Count
        ReDim arr(1 To cou)
        
        Dim i As Long
        i = LBound(arr)
        Dim sht As Worksheet
        For Each sht In ActiveWindow.selectedsheets '将被选择的sheet名称存储在数组中
            arr(i) = sht.Name
            i = i + 1
        Next
    
        For Each sht In Sheets
            If sht.Visible = xlSheetVisible Then '如果为显示状态
                Dim flag As Byte
                flag = False    '假设表格没有被选中，可删除
                For i = LBound(arr) To UBound(arr)
                    If sht.Name = arr(i) Then
                        flag = True '代表已经被选中，不能删除
                        Exit For
                    End If
                Next
                If flag = False Then
                    Excel.Application.DisplayAlerts = False '忽略弹窗警告
                    On Error Resume Next
                    sht.Delete
                    Excel.Application.DisplayAlerts = True
                End If
            End If
        Next
        
        Application.ScreenUpdating = False
    End If

End Sub
'多工作簿的表合并在当前文件中()
Sub workbooksMerge(control As IRibbonControl)
    
    Dim str()
    Dim i As Integer
    Dim Wb, wb1 As Workbook
    Dim sht, sht1 As Worksheet
    Dim shtName
    shtName = ActiveSheet.Name
    Set wb1 = ActiveWorkbook
    Set sht1 = ActiveSheet
    
    On Error Resume Next '加上这句防止用户点击取消发生的错误
    str = Application.GetOpenFilename("Excel数据文件,*.xls*", Title:="请选择要合并的文件", MultiSelect:=True)
    Application.ScreenUpdating = False
    
        For i = LBound(str) To UBound(str)
            Set Wb = Workbooks.Open(str(i))
            For Each sht In Wb.Sheets
                sht.Copy After:=wb1.Sheets(wb1.Sheets.Count)
                wb1.Sheets(wb1.Sheets.Count).Name = Split(Wb.Name, ".")(0) & "_" & sht.Name '文件名中请勿出现“.”否则会有误
            Next
            Wb.Close
        Next
    Application.ScreenUpdating = True
    Sheets(shtName).Select

End Sub

Sub worksheetsMerge(control As IRibbonControl)
'需要事先选中标题行的下一个最左的数据，类似冻结窗格
'需要删除多余sheet，避免区域无法自动识别
'需要新建表格并复制表头，且空出第一列以放置各sheet名称

    Dim newsht As Worksheet
    Dim arr() As String
    Dim i, con As Integer
    
    Dim msg As Byte
    msg = MsgBox("请确定已选择要合并的sheet", vbOKCancel)
    
    If msg <> 1 Then
        Exit Sub
    Else
        Application.ScreenUpdating = False
        con = ActiveWindow.selectedsheets.Count
        ReDim arr(1 To con)
        i = 1
        Dim sht As Worksheet
        For Each sht In ActiveWindow.selectedsheets
            arr(i) = sht.Name
            i = i + 1
        Next
    
        On Error Resume Next '防止点击取消发生错误
        Dim rng
        Set rng = Application.InputBox(" 请选择非标题的数据区域的最左上单元格", Type:=8)
        If rng = False Then
            MsgBox ("您未选择单元格，程序已结束")
            Exit Sub
        Else
            Excel.Application.DisplayAlerts = False
            Dim arow, acol
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
            [A1] = "工作簿_工作表"
            [A1].End(xlDown).Delete '删除多余的单元格
            [b1].Select '方便之后复制粘贴
            Sheets(arr(1)).Select
            Cells(1, acol).Select
            MsgBox "为防止自动复制遇到合并单元格出错，下面请您手动复制标题行"
            Excel.Application.DisplayAlerts = True
        End If
        Application.ScreenUpdating = True
    End If
    
End Sub

'按列拆分表格()
Sub createWorksheetsByColumn(control As IRibbonControl)
'选择的单元格不能为第1行

    Dim actsht As Worksheet
    Set actsht = ActiveSheet '存储当前活动的sheet名称
    
    If MsgBox("请确定以下内容：" & Chr(10) _
    & "1. 该Sheet为清单式表格(首行为标题行，首列不为空)" & Chr(10) _
    & "2. 已选择拆分的依据列中的某单元格", vbOKCancel) = vbOK Then

        Dim rng As Range
        Set rng = ActiveCell

        Dim i, col As Integer
    
        Excel.Application.DisplayAlerts = False
        With actsht
    
            '复制并删除重复值
            Dim colnum
            colnum = rng.Column '选择单元格的列号
            .Cells(2, colnum).Resize(.Cells(2, colnum).End(xlDown).Row + 1, 1).Copy Sheets.Add(before:=Sheets(1)).Range("a1") '复制到最左侧，此时该表为sheets(1)
            Sheets(1).Range("A1:A" & Sheets(1).[A1].End(xlDown).Row).removeDuplicates Columns:=1, Header:=xlNo  '去除重复值
            
            Dim rnum
            rnum = Sheets(1).[A1].End(xlDown).Row   '存储不重复值的个数
            '按名称新建表格于最后
            On Error GoTo 100
            For i = 1 To rnum
                Sheets.Add(After:=Sheets(Sheets.Count)).Name = Sheets(1).Range("a" & i)
            Next
            Sheets(1).Delete
    
            '复制对应内容到指定表格
            For i = Sheets.Count To Sheets.Count - rnum + 1 Step -1
                .Range(.[A1], .[A1].End(xlDown).End(xlToRight)).AutoFilter Field:=colnum, CRITERIA1:=Sheets(i).Name    '启动筛选
                .Range(.[A1], .[A1].End(xlDown).End(xlToRight)).Copy Sheets(i).[A1]
            Next
    
            '取消筛选
            .Select
            .Range(.[A1], .[A1].End(xlDown).End(xlToRight)).AutoFilter Field:=1
            Selection.AutoFilter
        End With
        Excel.Application.DisplayAlerts = True
        
        MsgBox "拆分完成"
    End If
    Exit Sub

100:

    MsgBox "已有相同名称的Sheet" & "，请删除后重试"

End Sub

'活动表另存为工作簿文件()
Sub toXlsx(control As IRibbonControl)

    Dim sht As Worksheet, filePath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "请选择要保存到哪个文件夹（此操作会覆盖重名文件）"
        .AllowMultiSelect = False   '禁止多选
        If .Show Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "未选择文件夹，操作已中止"
            Exit Sub
        End If
    End With
    If Right(filePath, 1) <> "\" Then filePath = filePath & "\"

    Dim wbFileFormat As Variant
    wbFileFormat = ActiveWorkbook.FileFormat    '获取当前文件的类型：xlxm、xls、xlsx……

    '关闭屏幕更新，隐藏文件保存的过程
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False  '有重名直接文件覆盖
    For Each sht In ActiveWindow.selectedsheets
        sht.Copy '该方法会直接复制到新建的工作簿中，即新工作簿文件为之后激活的窗口
        On Error Resume Next
        ActiveWorkbook.SaveAs fileName:=filePath & sht.Name, FileFormat:=wbFileFormat
        ActiveWorkbook.Close True
    Next
    MsgBox "处理完成！稍后将打开所在文件夹"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    '打开文件浏览器，默认最小化窗口改为正常且获取焦点
    Shell "explorer.exe /n, /e, " & filePath, vbNormalFocus

End Sub
'根据选区创建工作表
Sub createWorksheetsByRange(control As IRibbonControl)
    
    Dim rng As Range
    Dim cou, flag As Byte
    Dim i As Integer
    Dim actsht, actsel
    actsht = ActiveSheet.Name '保存当前活动sheet
    actsel = Selection.Cells.Count
    
    For Each rng In Selection
        flag = 0
        For i = 1 To Sheets.Count
            If Sheets(i).Name = (rng.Value & "") Or rng.Value & "" = "" Then
                flag = 1
                Exit For
            End If
        Next
        If flag = 0 Then
            Sheets.Add(After:=Sheets(Sheets.Count)).Name = rng.Value
        Else
            cou = cou + 1 '创建失败数目
        End If

    Next
    If cou <> 0 Then
        MsgBox "有 " & cou & " / " & actsel & " Sheet因名称重复/空值创建失败"
    Else
        MsgBox actsel & " 张 Sheet表创建成功"
    End If
    Sheets(actsht).Select

End Sub
'生成目录
Sub createDirectory(control As IRibbonControl)

    If MsgBox("请确认已选择需要生成目录的Sheet", vbOKCancel) = vbOK Then
        Dim shtarr(), cou As Long, i As Long, sht As Worksheet
        cou = ActiveWindow.selectedsheets.Count
        ReDim shtarr(1 To cou)  '将所选sheet存入数组备用
        i = 1   '设置从1开始存储
        For Each sht In ActiveWindow.selectedsheets
            Set shtarr(i) = sht
            i = i + 1
        Next

        On Error GoTo cancelChoose '防止点击取消发生错误
        Dim texttype As VbMsgBoxResult
        texttype = MsgBox("请选择要展示的文字，Yes为Sheet表名，No为单元格文字内容，Cancel为""Sheet1_B2""", vbYesNoCancel)
        Dim rng As Range
        Set rng = Application.InputBox("请选择你要链接到哪个单元格", Type:=8)
        If Err.Number = 13 Then GoTo cancelChoose    '如果点击了取消

        Dim textvalue As String, rngaddress As String  '显示的文字
        rngaddress = rng.Address(0, 0)  '获取相对地址如"B2"
        Dim mulusht As Worksheet
        Set mulusht = Sheets.Add(before:=Sheets(1)) '   生成新sheet以存放目录
        With mulusht
            On Error Resume Next
            .Name = "浅北表格自动生成目录"
            
            With .Range("C3")
                .Value = "目录"
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                With .Font
                    .Name = "Microsoft YaHei UI"
                    .Size = 20
                End With
            End With
            Range("D3").Value = "COUTENTS"
            
            '目录下方绿色横线
            With .Range("C3:D3")
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ThemeColor = 10
                    .TintAndShade = -0.249977111117893
                    .Weight = xlThick
                End With
            End With
            '调整列宽
            Columns("A:A").ColumnWidth = 2
            Columns("B:B").ColumnWidth = 2
            Columns("C:C").ColumnWidth = 7.4
            Columns("D:D").ColumnWidth = 24
            Columns("E:E").ColumnWidth = 2
            Columns("F:F").ColumnWidth = 2

            For i = LBound(shtarr) To UBound(shtarr)
                Dim textdis As String
                Select Case texttype
                Case 6, vbYes '   YES
                    textdis = shtarr(i).Name
                Case 7, vbNo 'NO
                    textdis = shtarr(i).Range(rngaddress).Value
                Case 2, vbCancel 'CANCEL
                    textdis = shtarr(i).Name & "_" & shtarr(i).Range(rngaddress).Value
                End Select

                .Hyperlinks.Add Anchor:=Range("C" & i + 4), Address:="", SubAddress:= _
                    shtarr(i).Name & "!" & rngaddress, ScreenTip:="由浅北表格助手创建", TextToDisplay:=textdis

               Range("B" & (i + 4) & ":E" & (i + 4)).RowHeight = 26.2

            Next
            
            '隐藏未使用单元格
            .Range(.Range("B" & (i + 6)).EntireRow, .Range("B" & (i + 6)).EntireRow.End(xlDown)).EntireRow.Hidden = True
            .Range(Columns("G:G"), .Columns("G:G").End(xlToRight)).EntireColumn.Hidden = True
            '隐藏编辑栏、网格线及行列号
            Application.DisplayFormulaBar = False
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.DisplayHeadings = False
            
            '具体目录所在区域
            With Range([c5], [c5].End(xlDown))
                With .Font
                    .Name = "微软雅黑"
                    .Size = 11
                    .ThemeColor = xlThemeColorDark2
                    .TintAndShade = -0.749992370372631
                    .Underline = xlUnderlineStyleNone
                End With
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
            End With
            
            '整个的背景色
            With .Range(["b2"], Range("E" & (i + 4))).Interior
                .Pattern = xlSolid
                .Color = 16119285
            End With

            .Activate   '激活到该工作表
        End With
    Else
        GoTo cancelChoose
    End If
    Exit Sub

cancelChoose:
    MsgBox ("已取消操作")
End Sub
'========================
'
'     邮 件 合 并
'
'========================
Sub dataSources(control As IRibbonControl)
    If hasWorkSheet("浅北来源表") = False Then
        If MsgBox("请确认该表第一行为标题行，且名称不重复，其他行为具体数据", vbOKCancel) = vbOK Then
            On Error GoTo cannotReName
            ActiveSheet.Name = "浅北来源表"
        End If
    Else
        MsgBox "已存在同名工作表，请检查后重试"
    End If
    Exit Sub

cannotReName:
    MsgBox "不能重命名工作表，可能是因为开启了工作簿保护"
End Sub
Sub setTemplate(control As IRibbonControl)
    If hasWorkSheet("浅北模板表") = False Then
        On Error GoTo cannotReName
        ActiveSheet.Name = "浅北模板表"
    Else
        MsgBox "已存在同名工作表，请检查后重试"
    End If
    Exit Sub
cannotReName:
    MsgBox "不能重命名工作表，可能是因为开启了工作簿保护"
End Sub
Sub routineFormat(control As IRibbonControl)
    Dim rng As Range
    For Each rng In Selection
        If rng <> "" Then rng.Value = ">" & rng.Value & "<"
    Next
End Sub
Sub textFormat(control As IRibbonControl)
    Dim rng As Range
    For Each rng In Selection
        If rng <> "" Then rng.Value = "》" & rng.Value & "《"
    Next
End Sub
Sub imgFormat(control As IRibbonControl)
    Dim rng As Range
    For Each rng In Selection
        If rng <> "" Then rng.Value = "》" & rng.Value & "<"
    Next
End Sub
'开始合并
Sub startMerge(control As IRibbonControl)

    '判断是否存在模板表及来源表

    If hasWorkSheet("浅北来源表") And hasWorkSheet("浅北模板表") Then

        On Error GoTo cancelSelect
        Worksheets("浅北模板表").Select
        Dim ergodic As Range  '必须使用变体类型，否则会报错，返回值为一个二维数组
        Set ergodic = Application.InputBox("请选择要遍历模板表的哪个区域", Type:=8)
        On Error GoTo 0

        '获取选区中符合规范的单元格的相对地址字符串
        Dim routineCell(), imgCell(), textCell() '一维数组，存放单元格在模板表的相对地址字符串
        Dim routineCol(), imgCol(), textCol() '一维数组，存放该单元格在来源表的列号

        '在选择的区域中遍历，找到对应地址字符串及来源列号
        Dim c As Range, i As Long, j As Long, k As Long, Co As Long
        For Each c In ergodic
            With Worksheets("浅北来源表")
                If Left(c.Value, 1) = ">" And Right(c.Value, 1) = "<" Then
                    ReDim Preserve routineCell(i), routineCol(i)
                    routineCell(i) = c.Address(0, 0)
                    For Co = 1 To .UsedRange.Columns.Count
                        If .Cells(1, Co).Value = Mid(c.Value, 2, Len(c.Value) - 2) Then
                            routineCol(i) = Co
                            Exit For
                        End If
                    Next
                    i = i + 1
                ElseIf Left(c.Value, 1) = "》" And Right(c.Value, 1) = "《" Then
                    ReDim Preserve textCell(j), textCol(j)
                    textCell(j) = c.Address(0, 0)
                    For Co = 1 To .UsedRange.Columns.Count
                        If .Cells(1, Co).Value = Mid(c.Value, 2, Len(c.Value) - 2) Then
                            textCol(j) = Co
                            Exit For
                        End If
                    Next
                    j = j + 1
                ElseIf Left(c.Value, 1) = "》" And Right(c.Value, 1) = "<" Then
                    ReDim Preserve imgCell(k), imgCol(k)
                    imgCell(k) = c.Address(0, 0)
                    For Co = 1 To .UsedRange.Columns.Count
                        If .Cells(1, Co).Value = Mid(c.Value, 2, Len(c.Value) - 2) Then
                            imgCol(k) = Co
                            Exit For
                        End If
                    Next
                    k = k + 1
                Else
                End If
            End With
        Next

        '重命名sheet的依据
        Worksheets("浅北来源表").Select
        On Error GoTo cancelSelect
        Dim byCol As Range
        Set byCol = Application.InputBox("请定位到命名依据列单元格？（该列不能有重复值）", Type:=8)

        Dim t
        t = Now   '记录现在时间
        Application.ScreenUpdating = False

        '复制表并修改里面的内容
        Dim ro As Long
        For ro = 2 To Worksheets("浅北来源表").UsedRange.Rows.Count
            With Worksheets("浅北来源表")
                Dim curSht As Worksheet
                On Error Resume Next
                Set curSht = Worksheets.Add(before:=Sheets(1))
                curSht.Name = .Cells(ro, byCol.Column).Value

                ergodic.Copy curSht.Range(ergodic.Address(0, 0))
                
                For i = LBound(routineCell) To UBound(routineCell)
                    If routineCol(i) Then
                        curSht.Range(routineCell(i)).Value = .Cells(ro, routineCol(i)).Value
                    End If
                Next
                For j = LBound(textCell) To UBound(textCell)
                    If textCol(j) Then
                        curSht.Range(textCell(j)).Value = .Cells(ro, textCol(j)).Value
                    End If
                Next
                For k = LBound(imgCell) To UBound(imgCell)
                    If imgCol(k) Then
                        curSht.Range(imgCell(k)).Value = .Cells(ro, imgCol(k)).Value
                        Dim shp As Shape
                        Set shp = ActiveSheet.Shapes.AddPicture(curSht.Range(imgCell(k)).Value, msoFalse, msoCTrue, curSht.Range(imgCell(k)).MergeArea.Left, curSht.Range(imgCell(k)).MergeArea.Top, curSht.Range(imgCell(k)).MergeArea.width, curSht.Range(imgCell(k)).MergeArea.height) '可匹配合并单元格
                        shp.Placement = xlMoveAndSize '随单元格大小和位置改变
                    End If
                Next
            End With

        Next

        Application.ScreenUpdating = True
        Worksheets("浅北来源表").Select

    Else
        MsgBox "请先设置来源表和模板表"
    End If
    Exit Sub
    
cancelSelect:
    MsgBox "已取消选择，操作已中止"
End Sub

'汇总表格
Sub worksheetsInOne(control As IRibbonControl)
    
    Dim newsht As Worksheet
    Dim arr()

    Dim i, j, k, con As Integer
    
    If MsgBox("请确保已选中需要合并的表格", vbOKCancel) = vbOK Then
        Dim rng As Range
        Set rng = Application.InputBox(" 复制各表格的哪个区域？", Type:=8)
        If Err Then Exit Sub

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

        Excel.Application.ScreenUpdating = False
    
        con = ActiveWindow.selectedsheets.Count
        ReDim arr(1 To con)
    
        i = 1
        Dim sht As Worksheet
        For Each sht In ActiveWindow.selectedsheets
            arr(i) = sht.Name
            i = i + 1
        Next
    
        Set newsht = Sheets.Add(before:=Sheets(1), Count:=1)
        Dim shtName
        shtName = newsht.Name
    
        i = 1
        j = 1
        For k = LBound(arr) To UBound(arr)
            Sheets(arr(k)).Range(rngaddress).Copy Sheets(shtName).Cells(i, j)
    
            If j <= (hangshu - 1) * colcount Then
                j = j + colcount
            Else
                i = i + rowcount
                j = 1
            End If
        Next
    
    End If
    
    Excel.Application.ScreenUpdating = True

End Sub
'文件批量重命名()
Sub renameFile(control As IRibbonControl)
    On Error Resume Next
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets("文件批量重命名").Copy
    MsgBox "表格模板已复制，点击上面按钮即可开始重命名/移动啦"
    sht.Activate
End Sub

Sub csvToXlsx(control As IRibbonControl)
    'csv文件不能超过1048576行，否则会出错
    If MsgBox("请确认csv文件不超过1048576行", vbOKCancel) <> vbOK Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim str()
    On Error Resume Next
    str = Application.GetOpenFilename("csv文件(*.csv),*.csv", Title:="请选择要转换的文件", MultiSelect:=True)
    
    Dim i As Integer
    For i = LBound(str) To UBound(str)
        Dim Wb As Workbook
        Set Wb = Workbooks.Open(str(i), ReadOnly:=True)
        '保存为默认工作簿+常规工作簿文件
        Wb.SaveAs Replace(str(i), ".csv", ""), IIf(Application.VERSION >= 12, xlWorkbookDefault, xlWorkbookNormal)
        Wb.Close
    Next
    
    Application.ScreenUpdating = True
End Sub
Sub xlamToXls(control As IRibbonControl)
    Dim strFile, Wb As Workbook
    strFile = Application.GetOpenFilename(FileFilter:="Micrsoft Excel文件(*.xlam), *.xlam")
    If strFile = False Then Exit Sub
    With Workbooks.Open(strFile)
        .IsAddin = False
        .SaveAs fileName:=Replace(strFile, "xlam", "xls"), FileFormat:=xlExcel8
        .Close
    End With
End Sub
'旧Excel文件破解()
Sub fileDecryption(control As IRibbonControl)
    On Error Resume Next
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
    Dim i
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
'========================
'
'    视 图 与 安 全
'
'========================

'删除出错名称()
Sub deleteErrorName(control As IRibbonControl)
    Dim n As Name
    For Each n In ActiveWorkbook.Names
        If InStr(n, "#REF!") Then n.Delete
    Next
End Sub
'设置可编辑区域()
Sub setEditableRange(control As IRibbonControl)

    On Error GoTo cannotProtect
    Selection.Locked = False
    Selection.FormulaHidden = False
    
    Dim str As String
    str = InputBox("请输入加密密码（可为空）")
    MsgBox "请记住你的工作表加密密码：“" & str & "”"
    ActiveSheet.Protect Password:=str, DrawingObjects:=True, Contents:=True, Scenarios:=True
    ActiveSheet.EnableSelection = xlUnlockedCells
    Exit Sub
    
cannotProtect:
    MsgBox "请先设置单元格锁定/撤销工作表保护！"
End Sub
'撤销工作表保护()
Sub unrestrictedEditableRange(control As IRibbonControl)
    On Error Resume Next
    ActiveSheet.Unprotect
    If Err Then
        MsgBox "您输入的密码不正确！"
        Exit Sub
    Else
        ActiveSheet.Cells.Locked = True '所有单元格锁定状态恢复默认
        MsgBox "已撤销工作表密码保护！"
    End If
End Sub
'保护工作簿结构()
Sub openWorkbookProtection(control As IRibbonControl)
    On Error GoTo cannotProtect
    ActiveWorkbook.Unprotect

    Dim str As String
    str = InputBox("请输入工作簿保护密码（可为空）")    '点击取消会视为空字符串

    ActiveWorkbook.Protect Password:=str, Structure:=True, Windows:=False
    MsgBox "请记住你的工作簿结构保护密码：“" & str & "”"
    Exit Sub

cannotProtect:
    MsgBox "请解除工作簿保护后重试"
End Sub

Sub closeWorkbookProtection(control As IRibbonControl)
    
    Dim str As String
    str = InputBox("请输入解锁工作簿结构保护密码")
    If Err Then Exit Sub
    On Error GoTo cannotUnprotect
    ActiveWorkbook.Unprotect str
    MsgBox "已解除工作簿结构保护！"
    Exit Sub

cannotUnprotect:
    MsgBox "您输入的密码不正确，请稍后再试"
End Sub
'========================
'
'       关  于
'
'========================
Sub aboutSoft(control As IRibbonControl)
    aboutForm.Show
    Call loadFun
End Sub

Sub warning(control As IRibbonControl)
    MsgBox "本工具为VBA编写，操作不可撤销，请确认风险！"
End Sub

Public Function getHttpJson(url As String)

    Dim xHttp As Object
    Set xHttp = CreateObject("Microsoft.XMLHTTP")

    xHttp.Open "GET", url, False
    xHttp.send

    getHttpJson = xHttp.responsetext

End Function
Public Function DoRegExp(sOrignText As String, sPattern As String) As String

    Dim oRegExp As Object
    Set oRegExp = CreateObject("VBScript.Regexp")
    With oRegExp
        .Global = True      '匹配所有的符合项
        .IgnoreCase = True  '不区分大小写
        .Pattern = sPattern '正则规则

        '判断是否可以找到匹配的字符，若可以则返回True
        If .test(sOrignText) Then
'           '对字符串执行正则查找，返回所有的查找值的集合，若未找到，则为空
            Dim oMatches As Object
            '定义匹配子字符串集合对象
            Dim oSubMatches As Object
            Dim oMatch As Object

            Set oMatches = .Execute(sOrignText)
            For Each oMatch In oMatches
                DoRegExp = oMatch.SubMatches(0)
            Next
        Else
            DoRegExp = ""
        End If
    End With

    Set oRegExp = Nothing
    Set oMatches = Nothing
End Function

