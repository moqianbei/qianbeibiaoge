Attribute VB_Name = "Fun"
Option Explicit 'ǿ�Ʊ�������
'Option Base 0 '�����±��0��ʼ

Function JVLOOKUP(Lookup_value As String, Table_array As Range, Offset_num As Integer)
Attribute JVLOOKUP.VB_Description = "���������������������ĵ�һ����Ԫ������ҵ����������Ҳ�ڡ�N����Ԫ���ֵ��δ�ҵ�����0"
Attribute JVLOOKUP.VB_ProcData.VB_Invoke_Func = " \n20"

    'ʹ��vba��find������ʵ�ֹ������match��index����
    'δ�ҵ��᷵������0
    
    Dim rng As Range
    Set rng = Table_array.Find(Lookup_value, , , xlWhole)
    If Not rng Is Nothing Then
        JVLOOKUP = rng.Offset(0, Offset_num).Value '�Ҳ�Ϊ�������Ϊ��
    End If

End Function
'�й�ʽ����12234445
Function JRANK(Number As Double, Ref As Range, Optional Order = 0)
Attribute JRANK.VB_Description = "���й�ʽ����������ĳ������һ�������������������ֵ�Ĵ�С����������������ͬ����ռ�����Ρ�\neg���в��е�2�����4���Զ���Ϊ��3����122345��"
Attribute JRANK.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error Resume Next
    Dim rng, i As Integer, Only As New Collection
    For Each rng In Ref
        '���rng ���ڳɼ������䵼�뵽 only �����У�����ȥ���ظ�ֵ
        If Order = False Then   'Ϊ 0 ����ԣ���������Խ������Խ��
            If rng > Number Then Only.Add rng, CStr(rng)
        Else            '��������Խ������Խ����
            If rng < Number Then Only.Add rng, CStr(rng)
        End If
    Next
    '������only�����ݸ�����1��ֵ����������Ϊ���ս��
    JRANK = Only.Count + 1
End Function
'��������
'Ӣʽ����=SUMPRODUCT((��������1=����1)* (��������2=����2)* (��������>����))
Function PM(igji As Double, quyu As Range, Optional criferia1_range As Range, Optional criferia1 As Range, Optional criferia2_range As Range = 1, Optional criferia2 As Range, Optional criferia3_range As Range = 1, Optional criferia3 As Range = 1)
    'Set zhipianyi = quyu.Offset(1, 0).Resize(quyu.Rows.Count + 1)
    PM.FormulaArray = " Sum((Frequency((criferia1_range = criferia1) * (criferia1_range = criferia2) * (criferia3_range = criferia3) * ֵ��������, ֵ��������) > 0) * (qy.Offset(-1, 0).Resize(qy.Rows.Count + 1) > igji)) + 1"
End Function
Function JHYPELINK(Ref)
Attribute JHYPELINK.VB_Description = "����ȡ���ӡ����ص�Ԫ�����õ����ӣ����δ�ҵ������ش���ֵ"
Attribute JHYPELINK.VB_ProcData.VB_Invoke_Func = " \n20"
'��ȡĳ��Ԫ����ڲ�����
    Application.Volatile True
    With Ref.Hyperlinks(1)
        JHYPELINK = IIf(.Address = "", .SubAddress, .Address)
    End With
End Function
Function JSHENFENZHENG(ID As String, Optional getType As Integer = 7)
Attribute JSHENFENZHENG.VB_Description = "�����֤��Ϣ����ȡ�й���½�������֤��������Ϣ"
Attribute JSHENFENZHENG.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim shenfenID As New IDCard
    JSHENFENZHENG = shenfenID.Info(ID, getType)
End Function

'1Ϊ�У�0ΪŮ��2���
Function JRANKNAME(Optional ByVal Sex As Byte = 2)
Attribute JRANKNAME.VB_Description = "���������������һ���������������"
Attribute JRANKNAME.VB_ProcData.VB_Invoke_Func = " \n20"

    Dim xingming As String, xingshi As String, nanxingming As String, nvxingming As String
    
    '�����ظ�����Ӱ������ʱ�䣬�Ƚ��ж�ȡ
    xingshi = DATA_xingshi
    nanxingming = DATA_nanxingming
    nvxingming = DATA_nvxingming
    
    If Sex = 1 Then
        xingming = Split(xingshi, ",")(Int(Rnd * CountX(xingshi, ",") + 1)) & _
        Split(nanxingming, ",")(Int(Rnd * CountX(nanxingming, ",") + 1))
    
    ElseIf Sex = 0 Then
        xingming = Split(xingshi, ",")(Int(Rnd * CountX(xingshi, ",") + 1)) & _
        Split(nvxingming, ",")(Int(Rnd * CountX(nvxingming, ",") + 1))
    Else
        xingming = Split(xingshi, ",")(Int(Rnd * CountX(xingshi, ",") + 1)) & _
        Split(nanxingming & nvxingming, ",")(Int(Rnd * CountX(DATA_nanxingming & nvxingming, ",") + 1))
    End If
    JRANKNAME = xingming
End Function

