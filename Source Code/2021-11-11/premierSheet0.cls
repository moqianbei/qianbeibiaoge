VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "premierSheet0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    '��ʾ�û�
    If MsgBox("�˲������ܻ����15�����ң���ȷ��Ҫ������", vbOKCancel) <> vbOK Then Exit Sub

    Dim t
    t = Timer

    Application.ScreenUpdating = False
    Excel.Application.DisplayAlerts = False '���Ե�������

    '���ԭ������������
    premierSheet1.Cells.Clear:    premierSheet2.Cells.Clear:    premierSheet3.Cells.Clear
    If premierSheet0.FilterMode Then premierSheet0.Cells.AutoFilter 'ȡ��ɸѡ״̬��������end�ķ����޷��õ��������ݣ�����ɸѡ��������
    '����������sheet�� 1�������� ���Ƶ�һ������sheet����ȥ���ظ�ֵ
    premierSheet0.Range(premierSheet0.Range("c2"), premierSheet0.Range("c2").End(xlDown)).Copy premierSheet1.Range("A1")
    premierSheet1.Range(premierSheet1.Range("A1"), premierSheet1.Range("A1").End(xlDown)).removeDuplicates Columns:=1, Header:=xlNo
    premierSheet1.Range(premierSheet1.Range("A1"), premierSheet1.Range("A1").End(xlDown)).Copy
    'ת�õ���������������
    premierSheet2.[A1].PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, Transpose:=True

    Dim i As Long, col As Long, j As Long, k As Long
    col = premierSheet2.UsedRange.Columns.Count 'һ��������
    
    Dim temp As Worksheet
    Set temp = Sheets.Add(before:=Sheets(1))    '��ʱ��������ȥ���ظ�ֵ
    '������ݶ���һ�����У�VBAȥ��ʱ���Զ���չѡ�����ķ���Դ���³������б���
    
    For i = 1 To col

        premierSheet0.UsedRange.AutoFilter Field:=3, CRITERIA1:=Sheet7.Cells(1, i).Value

        temp.Columns(1).Delete  '������ݣ���ֹ���ź�������
        premierSheet0.Range(premierSheet0.Range("D1").End(xlDown), premierSheet0.Range("D1").End(xlDown).End(xlDown)).Copy temp.[A1]
        temp.Range("A1:A" & temp.[A1].End(xlDown).Row).removeDuplicates Columns:=1, Header:=xlNo  'ȥ���ظ�ֵ
        temp.Range("A1:A" & temp.[A1].End(xlDown).Row).Copy premierSheet2.Cells(2, i)
        On Error Resume Next
        premierSheet2.Columns(i).SpecialCells(xlCellTypeBlanks).Delete  'ɾ����ֵ��Ԫ��
    Next

    If premierSheet0.FilterMode Then premierSheet0.Cells.AutoFilter '���ɸѡ״̬
    
    '�������ݵ���ʱ����ȥ�ز������ֵ������
    temp.Columns(Array(1, 2)).Delete    '���֮ǰ�����ݣ���ֹ����
    premierSheet0.Range("C1:D1", premierSheet0.Range("C1:D1").End(xlDown)).Copy temp.[A1]
    temp.Range("A1:B" & temp.[A1:B1].End(xlDown).Row).removeDuplicates Columns:=Array(1, 2), Header:=xlYes
    temp.Columns("B:B").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

    Dim allR As Variant, pre3 As Variant
    pre3 = temp.Range("A2:B" & temp.[A2:B2].End(xlDown).Row)
    allR = premierSheet0.UsedRange
    temp.Delete 'ɾ����ʱ�������ݾ��Ѵ�������

    For j = 1 To UBound(pre3)
        k = 2   '�ӵڶ��п�ʼд������
        premierSheet3.Cells(1, j).Value = pre3(j, 1) & "_" & pre3(j, 2)    '����д�롰�ӱ�ʡ_�е��С�����
        For i = 1 To UBound(allR)
            If allR(i, 3) = pre3(j, 1) And allR(i, 4) = pre3(j, 2) And allR(i, 5) <> "" Then
                premierSheet3.Cells(k, j).Value = allR(i, 5)
                k = k + 1
            End If
        Next
    Next

    Application.ScreenUpdating = True
    Excel.Application.DisplayAlerts = True
    MsgBox "�Ѿ��������ɣ���ʱ" & Timer - t & "��"  '���Զ�Ϊ�����¹�ʽ����"    13s

End Sub