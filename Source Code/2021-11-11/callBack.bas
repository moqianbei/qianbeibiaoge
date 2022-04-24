Attribute VB_Name = "callBack"
Option Explicit
'=======================
'
'     �����ڲ�����
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
'  �� Ԫ �� �� �� �� ʽ
'=======================

'��Ԫ����ʾΪ����KM
'------- https://club.excelhome.net/thread-1312477-1-1.html -------------------------------
'1����λ����1λС�����Զ����ʽӦ����Ϊ0!.0,"��"��0"."#,"��" ��0"."0,"��"
'2����λ����0λС��������ǧ�ַ����Զ����ʽӦ����Ϊ#,##0,,"��"%
'3����λ����1λС��������ǧ�ַ����Զ����ʽӦ����Ϊ#,##0.0,,"��"%
'4����λ����2λС��������ǧ�ַ����Զ����ʽӦ����Ϊ#,##0.00,,"��"%
'�������Ҫ��λ��ȥ����"��"������λ����Ӧ˫���ţ�

'����#,##0,,%�Զ����ʽ������˵����
'1.����#,##0,,%����꽫�������%ǰ����סAlt����ͬʱ�����ΰ�ѡС���̵�����1��0������ʹ��Ctrl+J��
'2.��ѡ�������趨Ϊ���Զ����С��������ġ�%��
'--------------------------------------------------------------------------------------------
Private Sub digitalFormat(numberscale As String, Optional accuracy As Byte = 2, Optional hasSeparator As Boolean = False)
    Select Case numberscale
        Case "k", "K"
            Select Case accuracy
                Case 0
                    Selection.NumberFormatLocal = "0,""K"""     '"��VBA����Ҫʹ��""ת��
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
            If hasSeparator Then    '��ǧ�ַ�
                Select Case accuracy
                    Case 0
                        Selection.NumberFormatLocal = "#,##0,,""��""" & Chr(10) & "%"   'chr(10)�ڵ�Ԫ���ڻس���Ctrl+J��
                    Case 1
                        Selection.NumberFormatLocal = "#,##0.0,,""��""" & Chr(10) & "%"
                    Case 2
                        Selection.NumberFormatLocal = "#,##0.00,,""��""" & Chr(10) & "%"
                    Case Else
                End Select

            Else                    '��ǧ�ַ�
                Select Case accuracy
                    Case 0
                        
                        Selection.NumberFormatLocal = "0,,""��""" & Chr(10) & "%"
                    Case 1
                        Selection.NumberFormatLocal = "0,,.0""��""" & Chr(10) & "%"
                    Case 2
                        Selection.NumberFormatLocal = "0,,.00""��""" & Chr(10) & "%"
                    Case Else
                End Select
            End If
        Case "HM", "hm"
            Selection.NumberFormatLocal = "0!.00,,""��"""
    End Select

    Selection.WrapText = True   '�Զ�����
    Selection.EntireRow.AutoFit '�Զ��и�

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
'  �� �� ת �� �� �� ʽ
'=======================

Sub dateAcross(control As IRibbonControl)
    Selection.NumberFormatLocal = "yyyy-mm-dd;@"
End Sub
Sub dateText(control As IRibbonControl)
    Selection.NumberFormatLocal = "yyyy��mm��dd��;@"
End Sub

'=======================
'  �� Ԫ �� �� �� �� ��
'=======================
Sub addPrefix(control As IRibbonControl)
    Dim prefix As String
    On Error Resume Next
    prefix = InputBox("������Ҫ��ӵ�ǰ׺��")
    If Err Then Exit Sub
    Dim rng As Range
    For Each rng In Selection
        rng.Value = prefix & rng.Value
    Next
End Sub
Sub addSuffix(control As IRibbonControl)
    Dim suffix As String
    On Error Resume Next
    suffix = InputBox("������Ҫ��ӵĺ�׺��")
    If Err Then Exit Sub
    Dim rng As Range
    For Each rng In Selection
        rng.Value = rng.Value & suffix
    Next
    rng.EntireColumn.AutoFit
End Sub
'תΪ����
Sub transferDate(control As IRibbonControl)

    Dim rng As Range
    Set rng = Selection
    Dim findArr() As Variant
    findArr = Array("��", "��", "~", "��")  'ע��û�С��ա�
    Dim i As Long
    For i = 1 To rng.Columns.Count
        'ʹ�÷��з�ʽת���ַ���
        Selection.Columns(i).TextToColumns DataType:=xlDelimited, FieldInfo:=Array(1, 5), TrailingMinusNumbers:=True
              'xlDelimitedʹ�÷ָ����ָ�
              'TextQualifier��ʹ�õ����š�˫���Ż��ǲ�ʹ��������Ϊ�ı��޶���
        Selection.Columns(i).AutoFit    '�Զ������п���ֹ����####����
    Next

    For i = LBound(findArr) To UBound(findArr)
        Selection.Replace What:=findArr(i), Replacement:="/", LookAt:=xlPart, _
            SearchOrder:=xlByRows, SearchFormat:=False, _
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Next
    Selection.Replace What:="��", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    Selection.NumberFormatLocal = "yyyy-mm-dd;@"

End Sub
'ǧ����KMת��������()
Sub transferNumber(control As IRibbonControl)

    If MsgBox("��ȷ����ѡ��ÿ����Ԫ��������������" & Chr(10) & _
      "    1.�������������KM���е�һ����λ���õ�λλ��ĩβ" & Chr(10) & _
      "    2.���ܳ����������ţ��硰,��$����", vbOKCancel) = vbOK Then

        Dim rng As Range
        For Each rng In Selection
            Dim str As String
            str = Right(rng.Value, 1)
    
            If UCase(str) = "K" Or str = "ǧ" Then
                rng.Value = VBA.Strings.Left(rng, VBA.Strings.Len(rng) - 1) * 1000
                Selection.NumberFormatLocal = "0.00,""K"""
            ElseIf UCase(str) = "M" Then
                rng.Value = VBA.Strings.Left(rng, VBA.Strings.Len(rng) - 1) * 1000000
                rng.NumberFormatLocal = "0.00,,""M"""
            ElseIf str = "��" Then
                rng.Value = VBA.Strings.Left(rng, VBA.Strings.Len(rng) - 1) * 100000000
                rng.NumberFormatLocal = "0!.00,,""��"""
            ElseIf UCase(str) = "W" Or str = "��" Then
                rng.Value = VBA.Strings.Left(rng, VBA.Strings.Len(rng) - 1) * 10000
                rng.NumberFormatLocal = "0,,.00""��""" & Chr(10) & "%"
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
'  �� Ԫ �� �� �� �� ֤
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
        .InputTitle = "���֤��"
        .InputMessage = "18λ�ַ���15λ�ַ�"
        .ErrorTitle = "¼�����֤����"
        .ErrorMessage = "������Ŀ��ܲ��������֤�淶������"
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
        .InputTitle = "�ֻ���"
        .ErrorTitle = "11λ�ַ����������ո��-�ָ�"
        .InputMessage = "�ֻ��Ų���ȷ"
        .ErrorMessage = "�������ֻ��Ų����Ϲ淶������"
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
        .InputTitle = "���п���"
        .ErrorTitle = "13~19λ�����ַ�"
        .InputMessage = "���п��Ų���ȷ"
        .ErrorMessage = "���������п��Ų����Ϲ淶������"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With

End Sub
Sub premierLevel(control As IRibbonControl)

    If Selection.Columns.Count > 1 Then
        MsgBox "��ѡ��Ҫ����һ������������"
    Else
        '���Ʊ�񵽵�ǰ������
        On Error Resume Next
        Application.DisplayAlerts = False   '��ֹ�Ƿ��滻��ǰ��Ի���
        If hasWorkSheet("�й������������Ʊ���ɾ��") = False Then ThisWorkbook.premierSheet0.Copy before:=ActiveWorkbook.Worksheets(1)
        If hasWorkSheet("�й�һ���������Ʊ���ɾ��") = False Then ThisWorkbook.premierSheet1.Copy before:=ActiveWorkbook.Worksheets(1)
        If hasWorkSheet("�й������������Ʊ���ɾ��") = False Then ThisWorkbook.premierSheet2.Copy before:=ActiveWorkbook.Worksheets(1)
        If hasWorkSheet("�й������������Ʊ���ɾ��") = False Then ThisWorkbook.premierSheet3.Copy before:=ActiveWorkbook.Worksheets(1)
        Application.DisplayAlerts = True
        On Error GoTo 0

        '��������
        Application.DisplayAlerts = False   '��ֹ�Ƿ��滻�Ի���
        ActiveWorkbook.Worksheets("�й������������Ʊ���ɾ��").Cells.CreateNames Top:=True, Left:=False, Bottom:=False, Right:=False   '��ӵ����ƹ�����
        ActiveWorkbook.Worksheets("�й������������Ʊ���ɾ��").Cells.CreateNames Top:=True, Left:=False, Bottom:=False, Right:=False
        Application.DisplayAlerts = True
    
        '����sheet������ǰ����ʾ
        ActiveWorkbook.Worksheets("�й�һ���������Ʊ���ɾ��").Visible = False
        ActiveWorkbook.Worksheets("�й������������Ʊ���ɾ��").Visible = False
        ActiveWorkbook.Worksheets("�й������������Ʊ���ɾ��").Visible = False
        
        '�������ֵ
        Dim n As Name
        For Each n In ActiveWorkbook.Names
            If InStr(n, "#REF!") Then n.Delete
        Next
    
        Dim firstRng As String, sRow As Long
        firstRng = Selection.Cells(1).Address(0, 0)
        
        Dim pre1 As Variant, fm As String
        '����r��1�Ķ�ά����
        pre1 = ActiveWorkbook.Worksheets("�й�һ���������Ʊ���ɾ��").Range(ActiveWorkbook.Worksheets("�й�һ���������Ʊ���ɾ��").[A1], ActiveWorkbook.Worksheets("�й�һ���������Ʊ���ɾ��").[A1].End(xlDown))
        Dim i As Long
        For i = LBound(pre1) To UBound(pre1) - 1
            fm = fm & pre1(i, 1) & ","
        Next
        fm = fm & pre1(UBound(pre1), 1)
    
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=fm
            .ignoreBlank = True '���Կ�ֵ
            .InCellDropdown = True  '�ṩ�����б�
            .InputTitle = "һ������"
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True       '��ʾҪ����ʲô��Ϣ
            .ShowError = False      '����ʾ������Ϣ
            .ignoreBlank = True
        End With
    
        Selection.Cells(1).Offset(0, 1).Resize(1, 2).NumberFormatLocal = "@"
        Selection.Cells(1).Offset(0, 1) = "=INDIRECT(" & firstRng & ")"
        Selection.Cells(1).Offset(0, 2) = "=INDIRECT(" & Selection.Cells(1).Address(0, 0) & "&""_""&" & Selection.Cells(1).Offset(0, 1).Address(0, 0) & ")"
    
        MsgBox "�����һ���������ã�" & Chr(10) & "��������ö��������������ݣ��ɷֱ�ѡ�У��������Ϸ���Ԫ���ֵ��Ϊ������֤-���еĹ�ʽ����"

    End If

End Sub
'���������֤
Sub cleanVerification(control As IRibbonControl)
    Selection.Validation.Delete
End Sub

'========================
'
'  �� Ԫ �� �� Ϣ �� ȡ
'
'========================
Sub readIdentityCard(control As IRibbonControl)
    If Selection.Count = 1 Then
        Dim v As String
        v = Selection.Value & ""
        Dim idmsg As New IDCard
        If idmsg.Info(v, 6) Like "*��*�淶*" Then
            MsgBox idmsg.Info(v, 6)
        Else
            MsgBox "��ѯ�������֤��Ϣ���£�" & Chr(10) _
                & Chr(10) & "���᣺ " & idmsg.Info(v, 1) _
                & Chr(10) & "���գ� " & Format(Chr(9) & idmsg.Info(v, 2), "YYYY-MM-DD;@") _
                & Chr(10) & "���䣺 " & idmsg.Info(v, 3) _
                & Chr(10) & "���ࣺ " & idmsg.Info(v, 4) _
                & Chr(10) & "������ " & idmsg.Info(v, 5) _
                & Chr(10) & "�Ա� " & idmsg.Info(v, 6) _
                & Chr(10) & "У�飺 " & idmsg.Info(v, 7) _
                & Chr(10)
        End If
    Else
        MsgBox "ֻ��ѡ��һ����Ԫ��"
    End If
End Sub

Sub readBankCard(control As IRibbonControl)
' ------ ʹ�����API: http://www.zhaotool.com/api/bank  ---------------
    If Selection.Count = 1 Then
        Dim v As String
        v = Selection.Value & ""
        Dim bankC As New BankCard
        MsgBox "��ѯ�������п���Ϣ���£�" & Chr(10) _
        & Chr(10) & "�����������ƣ� " & bankC.Info(v, 1) _
        & Chr(10) & "���п������ͣ� " & bankC.Info(v, 2) _
        & Chr(10) & "�������е绰�� " & bankC.Info(v, 3) _
        & Chr(10) & "У������ȷ�ԣ� " & bankC.Info(v, 4) _
        & Chr(10)
    Else
        MsgBox "ֻ��ѡ��һ����Ԫ��"
    End If
End Sub

Sub readPhoneNumber(control As IRibbonControl)

    If Selection.Count = 1 Then
        Dim p As New phoneCard
        MsgBox p.Info(Selection.Value & "")
    Else
        MsgBox "ֻ��ѡ��һ����Ԫ��"
    End If

End Sub
Sub readPostCode(control As IRibbonControl)

End Sub
'========================
'
'    ԭ λ �� ճ ��
'
'========================

'ԭλճ��Ϊֵ��Դ��ʽ()
Sub pasteValue(control As IRibbonControl)
'ֻ�趨λ��Ҫִ�е����򼴿�

    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False
    Application.CutCopyMode = False

End Sub
'ԭλճ��Ϊ��ʾ��ֵ()
Sub pasteDisplayValue(control As IRibbonControl)
'ֻ�趨λ��Ҫִ�е����򼴿�

    If MsgBox("�˲������ܺķ�һЩ��Դ�����޷��жϣ�ȷ�Ͻ�����", vbOKCancel) = vbOK Then
        Dim rng As Range
        For Each rng In Selection
            rng.Value = rng.Text
        Next
    End If

End Sub
'ճ���и��п��ʽ()
Sub pasteRowHColumnW(control As IRibbonControl)
'�ɿ�sheet����

    Dim rng1 As Range, rng2 As Range
    Dim i As Integer, j As Integer

    Set rng1 = Selection
    
    On Error Resume Next
    Set rng2 = Application.InputBox(" �ø�ʽӦ�����ĸ�����", Type:=8)
    If Err.Number = 424 Then GoTo cancelSelect
    On Error GoTo 0
    
    If rng1.Cells.Count = 1 Then           '1��N ��θ���
        Dim rng As Range
        For Each rng In rng2
            rng.ColumnWidth = rng1.ColumnWidth
            rng.RowHeight = rng1.RowHeight
        Next
    ElseIf rng2.Cells.Count = 1 Then        'N��1 ���θ���
        For i = 1 To rng1.Rows.Count
            For j = 1 To rng1.Columns.Count
                rng2.Offset(i - 1, j - 1).ColumnWidth = rng1.Cells(i, j).ColumnWidth
                rng2.Offset(i - 1, j - 1).RowHeight = rng1.Cells(i, j).RowHeight
            Next
        Next
    Else
        MsgBox "ֻ��1��N��N��1����NΪ��������"
    End If
    Exit Sub '��ǰ�������������д�����ʾ

cancelSelect:
    MsgBox ("��ȡ��ѡ��/���룬�����ѽ���"):
    Err.Clear
End Sub
'========================
'
'    ͼ Ƭ ͼ �� �� ��
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
'��������ͼƬ()
Sub insertPicture(control As IRibbonControl)

    If MsgBox("��ȷ������������·����D:\folder\abc.jpg��", vbOKCancel) = vbOK Then

        Dim rng1 As Range
        Set rng1 = Selection.Cells '�洢ѡ��

        Dim picdir As String
        picdir = InputBox("������Ҫƫ�Ƶ�λ�ã���Ӣ�Ķ��ŷָ�(����y��,����x��)")
        If picdir <> "" Then
            Dim errRng As Range, errNumber As Long  '�洢������λ�ĵ�Ԫ��
            
            Dim x As Double, y As Double
            x = Val(Split(picdir, ",")(0))
            If InStr(picdir, ",") <> 0 Then y = Val(Split(picdir, ",")(1))
            Dim rng As Range, rng2 As Range
            Dim shp As Shape
            For Each rng In rng1
                Set rng2 = rng.Offset(x, y)
                If rng.Value <> "" Then
                    On Error GoTo cannotFind
                    Set shp = ActiveSheet.Shapes.AddPicture(rng.Value, msoFalse, msoCTrue, rng2.MergeArea.Left, rng2.MergeArea.Top, rng2.MergeArea.width, rng2.MergeArea.height) '��ƥ��ϲ���Ԫ��
                    shp.Placement = xlMoveAndSize '�浥Ԫ���С��λ�øı�
                End If
            Next
            If errNumber > 0 Then
                MsgBox "��" & errNumber & "��ͼƬ�޷���ӣ���Ϊ��ѡ�д���Ԫ��"
                errRng.Select
            End If
        End If
        
    End If
    Exit Sub
cannotFind:
    If errNumber = 0 Then
        Set errRng = rng
    Else
        Set errRng = Union(errRng, rng) '��ѡ��������֮��ѡ���չʾ����
    End If
    errNumber = errNumber + 1
    Resume Next
End Sub
'�������ͼ��()
Sub deleteShape(control As IRibbonControl)
    Dim msg As VbMsgBoxResult
    msg = MsgBox("��ȷ�������sheet�н�ͼƬ��Y����������ͼ�Σ�N����", vbYesNoCancel)
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
'��λ��ѡ���Ŀ�ֵ��Ԫ��()
Sub locateNull(control As IRibbonControl)
    Selection.SpecialCells(xlCellTypeBlanks).Select
End Sub
'ת������
Sub convertRowsColumns(control As IRibbonControl)
    
    Dim arr, cot As Long

    arr = Selection.Cells
    cot = Selection.Cells.Count

    Dim rng As Range
    Set rng = Application.InputBox("ת�ú�ĵ�Ԫ���������ĵ�һ����Ԫ��", Type:=8)
    If Err Then
        GoTo 100
    Else
        Dim liehang As String

        liehang = InputBox("��������Ҫת������������������"",""�ָ���lie,[hang]��")
        If Err Then GoTo 100
        
        Dim lie As Double, hang As Double

        lie = Val(Split(liehang, ",")(0))
        If InStr(liehang, ",") <> 0 Then hang = Val(Split(liehang, ",")(1))
        If hang = 0 Then hang = IIf(cot / lie = Int(cot / lie), cot / lie, Int(cot / lie + 1))
        If lie = 0 Then lie = IIf(cot / hang = Int(cot / hang), cot / hang, Int(cot / hang + 1))

        Dim zixing As String
        zixing = InputBox("��������Ҫת������ʽ��N���к��У�Z���к���")
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
    MsgBox ("��ȡ��ѡ��/���룬�����ѽ���"):

End Sub
'ȥ��
Sub removeDuplicates(control As IRibbonControl)
    Selection.removeDuplicates Columns:=1, Header:=xlNo
End Sub
'========================
'
'    �� �� �� �� ��
'
'========================
'ȡ��ȫ��sheet����()
Sub showAllWorksheets(control As IRibbonControl)
    Dim sht As Worksheet
    For Each sht In Worksheets
        sht.Visible = xlSheetVisible
    Next
End Sub

'ѡ���1�����пɼ�sheet()
Sub selectWorksheets(control As IRibbonControl)
    
    'Worksheets.Select������ȡ��sheets(1)��ѡ��״̬
    
    Dim arr() '�洢����ʾ�Ĺ�������
    '��ѡ���һ��sheet
    Dim k As Long
    k = 0
    Dim sht As Worksheet
    For Each sht In Worksheets
        If sht.Visible = True Then
           k = k + 1
           ReDim Preserve arr(1 To k) '���¶��������С��������֮ǰ��ֵ
           arr(k) = sht.Name
        End If
    Next
    
    If UBound(arr) - LBound(arr) > 0 Then
        Sheets(arr(LBound(arr) + 1)).Select '��ѡ�пɼ��ĵڶ���sheet��Ҳ��ȡ���˵�һ��sheet��ѡ��״̬
        For Each sht In Worksheets
            Dim flag As Byte
            flag = False    '����û���ظ�ֵ
            Dim i As Long
            For i = LBound(arr) + 1 To UBound(arr) '�ų���һ�ſɼ�sheet
                If sht.Name = arr(i) Then
                    flag = True
                    Exit For
                End If
            Next
            If flag = True Then sht.Select Replace:=False   '��ѡ�����滻֮ǰ��ѡ��
        Next
    Else
        MsgBox "����һ�Ź�����ɼ�"
    End If

End Sub

Sub deleteNotSelectedWorksheets(control As IRibbonControl)
    
    If MsgBox("�˲���������ɾ��������������ʾ��sheet����ȷ��", vbOKCancel) = vbOK Then
    
        Application.ScreenUpdating = False

        Dim arr() As String
        
        Dim cou As Long
        cou = ActiveWindow.selectedsheets.Count
        ReDim arr(1 To cou)
        
        Dim i As Long
        i = LBound(arr)
        Dim sht As Worksheet
        For Each sht In ActiveWindow.selectedsheets '����ѡ���sheet���ƴ洢��������
            arr(i) = sht.Name
            i = i + 1
        Next
    
        For Each sht In Sheets
            If sht.Visible = xlSheetVisible Then '���Ϊ��ʾ״̬
                Dim flag As Byte
                flag = False    '������û�б�ѡ�У���ɾ��
                For i = LBound(arr) To UBound(arr)
                    If sht.Name = arr(i) Then
                        flag = True '�����Ѿ���ѡ�У�����ɾ��
                        Exit For
                    End If
                Next
                If flag = False Then
                    Excel.Application.DisplayAlerts = False '���Ե�������
                    On Error Resume Next
                    sht.Delete
                    Excel.Application.DisplayAlerts = True
                End If
            End If
        Next
        
        Application.ScreenUpdating = False
    End If

End Sub
'�๤�����ı�ϲ��ڵ�ǰ�ļ���()
Sub workbooksMerge(control As IRibbonControl)
    
    Dim str()
    Dim i As Integer
    Dim Wb, wb1 As Workbook
    Dim sht, sht1 As Worksheet
    Dim shtName
    shtName = ActiveSheet.Name
    Set wb1 = ActiveWorkbook
    Set sht1 = ActiveSheet
    
    On Error Resume Next '��������ֹ�û����ȡ�������Ĵ���
    str = Application.GetOpenFilename("Excel�����ļ�,*.xls*", Title:="��ѡ��Ҫ�ϲ����ļ�", MultiSelect:=True)
    Application.ScreenUpdating = False
    
        For i = LBound(str) To UBound(str)
            Set Wb = Workbooks.Open(str(i))
            For Each sht In Wb.Sheets
                sht.Copy After:=wb1.Sheets(wb1.Sheets.Count)
                wb1.Sheets(wb1.Sheets.Count).Name = Split(Wb.Name, ".")(0) & "_" & sht.Name '�ļ�����������֡�.�����������
            Next
            Wb.Close
        Next
    Application.ScreenUpdating = True
    Sheets(shtName).Select

End Sub

Sub worksheetsMerge(control As IRibbonControl)
'��Ҫ����ѡ�б����е���һ����������ݣ����ƶ��ᴰ��
'��Ҫɾ������sheet�����������޷��Զ�ʶ��
'��Ҫ�½���񲢸��Ʊ�ͷ���ҿճ���һ���Է��ø�sheet����

    Dim newsht As Worksheet
    Dim arr() As String
    Dim i, con As Integer
    
    Dim msg As Byte
    msg = MsgBox("��ȷ����ѡ��Ҫ�ϲ���sheet", vbOKCancel)
    
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
    
        On Error Resume Next '��ֹ���ȡ����������
        Dim rng
        Set rng = Application.InputBox(" ��ѡ��Ǳ������������������ϵ�Ԫ��", Type:=8)
        If rng = False Then
            MsgBox ("��δѡ��Ԫ�񣬳����ѽ���")
            Exit Sub
        Else
            Excel.Application.DisplayAlerts = False
            Dim arow, acol
            arow = rng.Row
            acol = rng.Column
    
            Set newsht = Sheets.Add(before:=Sheets(1), Count:=1) '��Ҫ���count����ΪĬ�ϻ������ѡ��sheet������
    
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
            [A1] = "������_������"
            [A1].End(xlDown).Delete 'ɾ������ĵ�Ԫ��
            [b1].Select '����֮����ճ��
            Sheets(arr(1)).Select
            Cells(1, acol).Select
            MsgBox "Ϊ��ֹ�Զ����������ϲ���Ԫ��������������ֶ����Ʊ�����"
            Excel.Application.DisplayAlerts = True
        End If
        Application.ScreenUpdating = True
    End If
    
End Sub

'���в�ֱ��()
Sub createWorksheetsByColumn(control As IRibbonControl)
'ѡ��ĵ�Ԫ����Ϊ��1��

    Dim actsht As Worksheet
    Set actsht = ActiveSheet '�洢��ǰ���sheet����
    
    If MsgBox("��ȷ���������ݣ�" & Chr(10) _
    & "1. ��SheetΪ�嵥ʽ���(����Ϊ�����У����в�Ϊ��)" & Chr(10) _
    & "2. ��ѡ���ֵ��������е�ĳ��Ԫ��", vbOKCancel) = vbOK Then

        Dim rng As Range
        Set rng = ActiveCell

        Dim i, col As Integer
    
        Excel.Application.DisplayAlerts = False
        With actsht
    
            '���Ʋ�ɾ���ظ�ֵ
            Dim colnum
            colnum = rng.Column 'ѡ��Ԫ����к�
            .Cells(2, colnum).Resize(.Cells(2, colnum).End(xlDown).Row + 1, 1).Copy Sheets.Add(before:=Sheets(1)).Range("a1") '���Ƶ�����࣬��ʱ�ñ�Ϊsheets(1)
            Sheets(1).Range("A1:A" & Sheets(1).[A1].End(xlDown).Row).removeDuplicates Columns:=1, Header:=xlNo  'ȥ���ظ�ֵ
            
            Dim rnum
            rnum = Sheets(1).[A1].End(xlDown).Row   '�洢���ظ�ֵ�ĸ���
            '�������½���������
            On Error GoTo 100
            For i = 1 To rnum
                Sheets.Add(After:=Sheets(Sheets.Count)).Name = Sheets(1).Range("a" & i)
            Next
            Sheets(1).Delete
    
            '���ƶ�Ӧ���ݵ�ָ�����
            For i = Sheets.Count To Sheets.Count - rnum + 1 Step -1
                .Range(.[A1], .[A1].End(xlDown).End(xlToRight)).AutoFilter Field:=colnum, CRITERIA1:=Sheets(i).Name    '����ɸѡ
                .Range(.[A1], .[A1].End(xlDown).End(xlToRight)).Copy Sheets(i).[A1]
            Next
    
            'ȡ��ɸѡ
            .Select
            .Range(.[A1], .[A1].End(xlDown).End(xlToRight)).AutoFilter Field:=1
            Selection.AutoFilter
        End With
        Excel.Application.DisplayAlerts = True
        
        MsgBox "������"
    End If
    Exit Sub

100:

    MsgBox "������ͬ���Ƶ�Sheet" & "����ɾ��������"

End Sub

'������Ϊ�������ļ�()
Sub toXlsx(control As IRibbonControl)

    Dim sht As Worksheet, filePath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "��ѡ��Ҫ���浽�ĸ��ļ��У��˲����Ḳ�������ļ���"
        .AllowMultiSelect = False   '��ֹ��ѡ
        If .Show Then
            filePath = .SelectedItems(1)
        Else
            MsgBox "δѡ���ļ��У���������ֹ"
            Exit Sub
        End If
    End With
    If Right(filePath, 1) <> "\" Then filePath = filePath & "\"

    Dim wbFileFormat As Variant
    wbFileFormat = ActiveWorkbook.FileFormat    '��ȡ��ǰ�ļ������ͣ�xlxm��xls��xlsx����

    '�ر���Ļ���£������ļ�����Ĺ���
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False  '������ֱ���ļ�����
    For Each sht In ActiveWindow.selectedsheets
        sht.Copy '�÷�����ֱ�Ӹ��Ƶ��½��Ĺ������У����¹������ļ�Ϊ֮�󼤻�Ĵ���
        On Error Resume Next
        ActiveWorkbook.SaveAs fileName:=filePath & sht.Name, FileFormat:=wbFileFormat
        ActiveWorkbook.Close True
    Next
    MsgBox "������ɣ��Ժ󽫴������ļ���"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    '���ļ��������Ĭ����С�����ڸ�Ϊ�����һ�ȡ����
    Shell "explorer.exe /n, /e, " & filePath, vbNormalFocus

End Sub
'����ѡ������������
Sub createWorksheetsByRange(control As IRibbonControl)
    
    Dim rng As Range
    Dim cou, flag As Byte
    Dim i As Integer
    Dim actsht, actsel
    actsht = ActiveSheet.Name '���浱ǰ�sheet
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
            cou = cou + 1 '����ʧ����Ŀ
        End If

    Next
    If cou <> 0 Then
        MsgBox "�� " & cou & " / " & actsel & " Sheet�������ظ�/��ֵ����ʧ��"
    Else
        MsgBox actsel & " �� Sheet�����ɹ�"
    End If
    Sheets(actsht).Select

End Sub
'����Ŀ¼
Sub createDirectory(control As IRibbonControl)

    If MsgBox("��ȷ����ѡ����Ҫ����Ŀ¼��Sheet", vbOKCancel) = vbOK Then
        Dim shtarr(), cou As Long, i As Long, sht As Worksheet
        cou = ActiveWindow.selectedsheets.Count
        ReDim shtarr(1 To cou)  '����ѡsheet�������鱸��
        i = 1   '���ô�1��ʼ�洢
        For Each sht In ActiveWindow.selectedsheets
            Set shtarr(i) = sht
            i = i + 1
        Next

        On Error GoTo cancelChoose '��ֹ���ȡ����������
        Dim texttype As VbMsgBoxResult
        texttype = MsgBox("��ѡ��Ҫչʾ�����֣�YesΪSheet������NoΪ��Ԫ���������ݣ�CancelΪ""Sheet1_B2""", vbYesNoCancel)
        Dim rng As Range
        Set rng = Application.InputBox("��ѡ����Ҫ���ӵ��ĸ���Ԫ��", Type:=8)
        If Err.Number = 13 Then GoTo cancelChoose    '��������ȡ��

        Dim textvalue As String, rngaddress As String  '��ʾ������
        rngaddress = rng.Address(0, 0)  '��ȡ��Ե�ַ��"B2"
        Dim mulusht As Worksheet
        Set mulusht = Sheets.Add(before:=Sheets(1)) '   ������sheet�Դ��Ŀ¼
        With mulusht
            On Error Resume Next
            .Name = "ǳ������Զ�����Ŀ¼"
            
            With .Range("C3")
                .Value = "Ŀ¼"
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                With .Font
                    .Name = "Microsoft YaHei UI"
                    .Size = 20
                End With
            End With
            Range("D3").Value = "COUTENTS"
            
            'Ŀ¼�·���ɫ����
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
            '�����п�
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
                    shtarr(i).Name & "!" & rngaddress, ScreenTip:="��ǳ��������ִ���", TextToDisplay:=textdis

               Range("B" & (i + 4) & ":E" & (i + 4)).RowHeight = 26.2

            Next
            
            '����δʹ�õ�Ԫ��
            .Range(.Range("B" & (i + 6)).EntireRow, .Range("B" & (i + 6)).EntireRow.End(xlDown)).EntireRow.Hidden = True
            .Range(Columns("G:G"), .Columns("G:G").End(xlToRight)).EntireColumn.Hidden = True
            '���ر༭���������߼����к�
            Application.DisplayFormulaBar = False
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.DisplayHeadings = False
            
            '����Ŀ¼��������
            With Range([c5], [c5].End(xlDown))
                With .Font
                    .Name = "΢���ź�"
                    .Size = 11
                    .ThemeColor = xlThemeColorDark2
                    .TintAndShade = -0.749992370372631
                    .Underline = xlUnderlineStyleNone
                End With
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
            End With
            
            '�����ı���ɫ
            With .Range(["b2"], Range("E" & (i + 4))).Interior
                .Pattern = xlSolid
                .Color = 16119285
            End With

            .Activate   '����ù�����
        End With
    Else
        GoTo cancelChoose
    End If
    Exit Sub

cancelChoose:
    MsgBox ("��ȡ������")
End Sub
'========================
'
'     �� �� �� ��
'
'========================
Sub dataSources(control As IRibbonControl)
    If hasWorkSheet("ǳ����Դ��") = False Then
        If MsgBox("��ȷ�ϸñ��һ��Ϊ�����У������Ʋ��ظ���������Ϊ��������", vbOKCancel) = vbOK Then
            On Error GoTo cannotReName
            ActiveSheet.Name = "ǳ����Դ��"
        End If
    Else
        MsgBox "�Ѵ���ͬ�����������������"
    End If
    Exit Sub

cannotReName:
    MsgBox "������������������������Ϊ�����˹���������"
End Sub
Sub setTemplate(control As IRibbonControl)
    If hasWorkSheet("ǳ��ģ���") = False Then
        On Error GoTo cannotReName
        ActiveSheet.Name = "ǳ��ģ���"
    Else
        MsgBox "�Ѵ���ͬ�����������������"
    End If
    Exit Sub
cannotReName:
    MsgBox "������������������������Ϊ�����˹���������"
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
        If rng <> "" Then rng.Value = "��" & rng.Value & "��"
    Next
End Sub
Sub imgFormat(control As IRibbonControl)
    Dim rng As Range
    For Each rng In Selection
        If rng <> "" Then rng.Value = "��" & rng.Value & "<"
    Next
End Sub
'��ʼ�ϲ�
Sub startMerge(control As IRibbonControl)

    '�ж��Ƿ����ģ�����Դ��

    If hasWorkSheet("ǳ����Դ��") And hasWorkSheet("ǳ��ģ���") Then

        On Error GoTo cancelSelect
        Worksheets("ǳ��ģ���").Select
        Dim ergodic As Range  '����ʹ�ñ������ͣ�����ᱨ������ֵΪһ����ά����
        Set ergodic = Application.InputBox("��ѡ��Ҫ����ģ�����ĸ�����", Type:=8)
        On Error GoTo 0

        '��ȡѡ���з��Ϲ淶�ĵ�Ԫ�����Ե�ַ�ַ���
        Dim routineCell(), imgCell(), textCell() 'һά���飬��ŵ�Ԫ����ģ������Ե�ַ�ַ���
        Dim routineCol(), imgCol(), textCol() 'һά���飬��Ÿõ�Ԫ������Դ����к�

        '��ѡ��������б������ҵ���Ӧ��ַ�ַ�������Դ�к�
        Dim c As Range, i As Long, j As Long, k As Long, Co As Long
        For Each c In ergodic
            With Worksheets("ǳ����Դ��")
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
                ElseIf Left(c.Value, 1) = "��" And Right(c.Value, 1) = "��" Then
                    ReDim Preserve textCell(j), textCol(j)
                    textCell(j) = c.Address(0, 0)
                    For Co = 1 To .UsedRange.Columns.Count
                        If .Cells(1, Co).Value = Mid(c.Value, 2, Len(c.Value) - 2) Then
                            textCol(j) = Co
                            Exit For
                        End If
                    Next
                    j = j + 1
                ElseIf Left(c.Value, 1) = "��" And Right(c.Value, 1) = "<" Then
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

        '������sheet������
        Worksheets("ǳ����Դ��").Select
        On Error GoTo cancelSelect
        Dim byCol As Range
        Set byCol = Application.InputBox("�붨λ�����������е�Ԫ�񣿣����в������ظ�ֵ��", Type:=8)

        Dim t
        t = Now   '��¼����ʱ��
        Application.ScreenUpdating = False

        '���Ʊ��޸����������
        Dim ro As Long
        For ro = 2 To Worksheets("ǳ����Դ��").UsedRange.Rows.Count
            With Worksheets("ǳ����Դ��")
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
                        Set shp = ActiveSheet.Shapes.AddPicture(curSht.Range(imgCell(k)).Value, msoFalse, msoCTrue, curSht.Range(imgCell(k)).MergeArea.Left, curSht.Range(imgCell(k)).MergeArea.Top, curSht.Range(imgCell(k)).MergeArea.width, curSht.Range(imgCell(k)).MergeArea.height) '��ƥ��ϲ���Ԫ��
                        shp.Placement = xlMoveAndSize '�浥Ԫ���С��λ�øı�
                    End If
                Next
            End With

        Next

        Application.ScreenUpdating = True
        Worksheets("ǳ����Դ��").Select

    Else
        MsgBox "����������Դ���ģ���"
    End If
    Exit Sub
    
cancelSelect:
    MsgBox "��ȡ��ѡ�񣬲�������ֹ"
End Sub

'���ܱ��
Sub worksheetsInOne(control As IRibbonControl)
    
    Dim newsht As Worksheet
    Dim arr()

    Dim i, j, k, con As Integer
    
    If MsgBox("��ȷ����ѡ����Ҫ�ϲ��ı��", vbOKCancel) = vbOK Then
        Dim rng As Range
        Set rng = Application.InputBox(" ���Ƹ������ĸ�����", Type:=8)
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
        hangshu = InputBox("������ÿ�зŶ��ٸ����")

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
'�ļ�����������()
Sub renameFile(control As IRibbonControl)
    On Error Resume Next
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets("�ļ�����������").Copy
    MsgBox "���ģ���Ѹ��ƣ�������水ť���ɿ�ʼ������/�ƶ���"
    sht.Activate
End Sub

Sub csvToXlsx(control As IRibbonControl)
    'csv�ļ����ܳ���1048576�У���������
    If MsgBox("��ȷ��csv�ļ�������1048576��", vbOKCancel) <> vbOK Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Dim str()
    On Error Resume Next
    str = Application.GetOpenFilename("csv�ļ�(*.csv),*.csv", Title:="��ѡ��Ҫת�����ļ�", MultiSelect:=True)
    
    Dim i As Integer
    For i = LBound(str) To UBound(str)
        Dim Wb As Workbook
        Set Wb = Workbooks.Open(str(i), ReadOnly:=True)
        '����ΪĬ�Ϲ�����+���湤�����ļ�
        Wb.SaveAs Replace(str(i), ".csv", ""), IIf(Application.VERSION >= 12, xlWorkbookDefault, xlWorkbookNormal)
        Wb.Close
    Next
    
    Application.ScreenUpdating = True
End Sub
Sub xlamToXls(control As IRibbonControl)
    Dim strFile, Wb As Workbook
    strFile = Application.GetOpenFilename(FileFilter:="Micrsoft Excel�ļ�(*.xlam), *.xlam")
    If strFile = False Then Exit Sub
    With Workbooks.Open(strFile)
        .IsAddin = False
        .SaveAs fileName:=Replace(strFile, "xlam", "xls"), FileFormat:=xlExcel8
        .Close
    End With
End Sub
'��Excel�ļ��ƽ�()
Sub fileDecryption(control As IRibbonControl)
    On Error Resume Next
    Dim fileName
    fileName = Application.GetOpenFilename("Excel�ļ�,*.xls;*.xla;*.xlt", , "�ɸ�ʽ�ļ��ƽ�")
    If Dir(fileName) = "" Then
        MsgBox "û�ҵ�����ļ�,����������"
        Exit Sub
    Else
        FileCopy fileName, fileName & ".bak" '�����ļ�
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
        MsgBox "���ȶ�VBA��������һ����������...", 32, "��ʾ"
        Exit Sub
    End If
    
    Dim St As String * 2
    Dim s20 As String * 1
    'ȡ��һ��0D0Aʮ�������ִ�
    Get #1, CMGs - 2, St
    'ȡ��һ��20ʮ�����ִ�
    Get #1, DPBo + 16, s20
    '�滻���ܲ��ݻ���
    For i = CMGs To DPBo Step 2
        Put #1, i, St
    Next
    '���벻��Է���
    If (DPBo - CMGs) Mod 2 <> 0 Then
        Put #1, DPBo + 1, s20
    End If
    MsgBox "�ļ����ܳɹ�......", 32, "��ʾ"
    Close #1
End Sub
'========================
'
'    �� ͼ �� �� ȫ
'
'========================

'ɾ����������()
Sub deleteErrorName(control As IRibbonControl)
    Dim n As Name
    For Each n In ActiveWorkbook.Names
        If InStr(n, "#REF!") Then n.Delete
    Next
End Sub
'���ÿɱ༭����()
Sub setEditableRange(control As IRibbonControl)

    On Error GoTo cannotProtect
    Selection.Locked = False
    Selection.FormulaHidden = False
    
    Dim str As String
    str = InputBox("������������루��Ϊ�գ�")
    MsgBox "���ס��Ĺ�����������룺��" & str & "��"
    ActiveSheet.Protect Password:=str, DrawingObjects:=True, Contents:=True, Scenarios:=True
    ActiveSheet.EnableSelection = xlUnlockedCells
    Exit Sub
    
cannotProtect:
    MsgBox "�������õ�Ԫ������/��������������"
End Sub
'������������()
Sub unrestrictedEditableRange(control As IRibbonControl)
    On Error Resume Next
    ActiveSheet.Unprotect
    If Err Then
        MsgBox "����������벻��ȷ��"
        Exit Sub
    Else
        ActiveSheet.Cells.Locked = True '���е�Ԫ������״̬�ָ�Ĭ��
        MsgBox "�ѳ������������뱣����"
    End If
End Sub
'�����������ṹ()
Sub openWorkbookProtection(control As IRibbonControl)
    On Error GoTo cannotProtect
    ActiveWorkbook.Unprotect

    Dim str As String
    str = InputBox("�����빤�����������루��Ϊ�գ�")    '���ȡ������Ϊ���ַ���

    ActiveWorkbook.Protect Password:=str, Structure:=True, Windows:=False
    MsgBox "���ס��Ĺ������ṹ�������룺��" & str & "��"
    Exit Sub

cannotProtect:
    MsgBox "��������������������"
End Sub

Sub closeWorkbookProtection(control As IRibbonControl)
    
    Dim str As String
    str = InputBox("����������������ṹ��������")
    If Err Then Exit Sub
    On Error GoTo cannotUnprotect
    ActiveWorkbook.Unprotect str
    MsgBox "�ѽ���������ṹ������"
    Exit Sub

cannotUnprotect:
    MsgBox "����������벻��ȷ�����Ժ�����"
End Sub
'========================
'
'       ��  ��
'
'========================
Sub aboutSoft(control As IRibbonControl)
    aboutForm.Show
    Call loadFun
End Sub

Sub warning(control As IRibbonControl)
    MsgBox "������ΪVBA��д���������ɳ�������ȷ�Ϸ��գ�"
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
        .Global = True      'ƥ�����еķ�����
        .IgnoreCase = True  '�����ִ�Сд
        .Pattern = sPattern '�������

        '�ж��Ƿ�����ҵ�ƥ����ַ����������򷵻�True
        If .test(sOrignText) Then
'           '���ַ���ִ��������ң��������еĲ���ֵ�ļ��ϣ���δ�ҵ�����Ϊ��
            Dim oMatches As Object
            '����ƥ�����ַ������϶���
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

