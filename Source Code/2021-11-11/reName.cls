VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "�ļ�����������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButtonReChoose_Click()

With ActiveWorkbook.Sheets("�ļ�����������")
    Dim str
    On Error Resume Next '��������ֹ�û����ȡ�������Ĵ���
    str = Application.GetOpenFilename("�����ļ�,*.*,PDF�ļ�,*.pdf;*.PDF,Word�ĵ�,*.doc*,Excel����,*.xls*", FilterIndex:=1, Title:="��ѡ��Ҫ���������ļ�", MultiSelect:=True)
    If Err Then MsgBox "��ȡ����������������ֹ": Exit Sub
    Application.ScreenUpdating = False

    .Range(.Range("A4:h4"), .Range("A4:h4").End(xlDown)).Clear
    .Range("a3") = "M:\ǳ����������\ʾ���ĵ�.xlsm"

    For i = LBound(str) To UBound(str)
        .Range("a" & i + 3) = str(i)    '�ӵ�4�е�Ԫ��ʼ��ȡѡ����ļ�
    Next
    '�Զ���乫ʽ
    .Range("B3:H3").AutoFill destination:=.Range("B3:H" & UBound(str) + 3), Type:=xlFunction

    Application.ScreenUpdating = True

    MsgBox "�༭��ɺ����Ϸ���ʼ��ť����"
    .Select
End With
      
End Sub

Private Sub CommandButtonRename_Click()
' ����������

    If Range("a3") = "M:\ǳ����������\ʾ���ĵ�.xlsm" Then
        MsgBox "δѡ���κ��ļ����������˳�"
        Exit Sub
    Else
        t = Timer
        With ActiveWorkbook.Sheets("�ļ�����������")
            For i = 3 To Range("a3").End(xlDown).Row
                On Error Resume Next 'Ϊ��ֹ�����ļ�
                If Range("A" & i).Value = "" Then Exit For
                Name Range("a" & i) As Range("h" & i)
            Next
        End With
        MsgBox "���������������������ʱ" & Timer - t & "�룡"
    End If
End Sub
