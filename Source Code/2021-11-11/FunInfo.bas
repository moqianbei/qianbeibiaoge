Attribute VB_Name = "��������"
Option Explicit
'�������ܽ��ܣ����ܷ��ڳ�������ʱ���أ���Ҫ�����ť����
Sub loadFun()

    Dim JVLOOKUPArg(1 To 3) As String
    JVLOOKUPArg(1) = "������ֵ����Ҫ�����ݱ�����������ֵ����������ֵ�����û��ַ���"
    JVLOOKUPArg(2) = "������������������"
    JVLOOKUPArg(3) = "�����λ�á�����ҵ��˸����ݣ�Ҫ�������Ҳ�ĵڼ�����Ԫ������ݡ�0Ϊ���ݱ�����������ʾ�������"
    Excel.Application.MacroOptions "JVLOOKUP", _
    "���������������������ĵ�һ����Ԫ������ҵ����������Ҳ�ڡ�N����Ԫ���ֵ��δ�ҵ�����0", _
    Category:="ǳ����������", _
    ArgumentDescriptions:=JVLOOKUPArg

'-------------------------------------------------------------------------------------------------

    Dim JRANKArg(1 To 3) As String
    JRANKArg(1) = "�����֡����ĸ����ֵ�����"
    JRANKArg(2) = "��������������ĸ����������"
    JRANKArg(3) = "[����ʽ] ָ�������ķ�ʽ��0����ԣ�����ֵԽ������Խ��ǰ����1:��֮"
    Excel.Application.MacroOptions "JRANK", _
    "���й�ʽ����������ĳ������һ�������������������ֵ�Ĵ�С����������������ͬ����ռ�����Ρ�" & Chr(10) & "eg���в��е�2�����4���Զ���Ϊ��3����122345��", _
    Category:="ǳ����������", _
    ArgumentDescriptions:=JRANKArg

'-------------------------------------------------------------------------------------------------

    Dim JSHENFENZHENGArg(1 To 2) As String
    JSHENFENZHENGArg(1) = "������֤�š�15λ��18λ�ַ�"
    JSHENFENZHENGArg(2) = "[2��Ϣ����] ��ȡ����֤�е�������Ϣ" & Chr(10) & _
      "1:����  2:����  3:����  4:��Ф  5:����  6:�Ա�" & Chr(10) & _
      "7:�Ƿ�Ϲ棨Ĭ�ϣ�  8:У����  9:ת18λ����"
    Excel.Application.MacroOptions "JSHENFENZHENG", _
    "������֤��Ϣ����ȡ�й���½��������֤��������Ϣ", _
    Category:="ǳ����������", _
    ArgumentDescriptions:=JSHENFENZHENGArg

'-------------------------------------------------------------------------------------------------

    Excel.Application.MacroOptions "JHYPELINK", _
    "����ȡ���ӡ����ص�Ԫ�����õ����ӣ����δ�ҵ������ش���ֵ", _
    Category:="ǳ����������", _
    ArgumentDescriptions:="��Ҫ��ȡ���ӵĵ�Ԫ��"

'-------------------------------------------------------------------------------------------------

    Excel.Application.MacroOptions "JRANKNAME", _
    "���������������һ���������������", _
    Category:="ǳ����������", _
    ArgumentDescriptions:="[�Ա�] 1:��  0:Ů  2:�����Ĭ��ֵ��"

End Sub