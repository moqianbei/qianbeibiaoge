VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'��ȡ����֤��Ϣ
Property Get Info(ID As String, Optional getType As Integer = 7) As Variant
'ID��Ϊ15λ��18λ�����ؿ�
'1Ϊ������δ�ҵ����ؿ�
'2Ϊ�����������ڣ���������ִ��󷵻أ�2������ֵ�����ش�����ʾ�ַ���
    '3Ϊ���䣨�������֣�
    '4Ϊ��Ф
    '5Ϊ����
'6Ϊ�Ա���Ů
'7�ж�У�����Ƿ���ȷ������ֵBoolean

    Info = "����֤�Ų����Ϲ淶"   '�趨Ĭ�Ϸ���ֵ
    
    If Len(ID) = 15 Then
        Info = Info(ToRule(ID), getType)
    ElseIf Len(ID) = 18 Then

        Dim errDataMsg As String, errPlaceMsg As String
        errDataMsg = "����֤������������"
        errPlaceMsg = "�Ҳ�������֤���ڵ���"

        Select Case getType
            Case 1    '1Ϊ����
                Dim pre6 As String
                pre6 = Left(ID, 6)
                On Error Resume Next
                '��Ԫ��ƥ��
                Info = premierSheet0.Range("A:A").Find(What:=pre6, LookAt:=xlWhole).Offset(0, 1).Value
                If Err.Number = 91 Then
                    Info = errPlaceMsg
                    Err.Clear
                End If
                On Error GoTo 0

            Case 2  '��������ת��Ϊ��������
                If Format(DateSerial(Mid(ID, 7, 4), Mid(ID, 11, 2), Mid(ID, 13, 2)), "yyyymmdd") = Mid(ID, 7, 8) Then
                    Info = DateSerial(Mid(ID, 7, 4), Mid(ID, 11, 2), Mid(ID, 13, 2))
                Else
                    Info = errDataMsg
                End If
    
            Case 3    '3Ϊ����
                If Info(ID, 2) <> errDataMsg Then
                    Info = DateDiff("yyyy", Info(ID, 2), Date)
                Else
                    Info = errDataMsg
                End If
            Case 4    '4Ϊ��Ф
                Dim shengxiao() As Variant
                shengxiao = Array("��", "ţ", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��")

                If Info(ID, 2) <> errDataMsg Then
                    Info = shengxiao((Mid(ID, 7, 4) - 1900) Mod 12) '1900��Ϊ����
                Else
                    Info = errDataMsg
                End If
            Case 5    '5Ϊ����
                If Info(ID, 2) <> errDataMsg Then
                    Dim xingzuo
                    xingzuo = [{"ˮƿ��",0121,0219;"˫����",0220,0320;"������",0321,0420;"��ţ��",0421,0521;"˫����",0522,0621;"��з��",0622,0723;"ʨ����",0724,0823;"��Ů��",0824,0923;"�����",0924,1023;"��Ы��",1024,1122;"������",1123,1222;"Ħ����",1223,0122}]
                    Dim id4
                    id4 = Val(Mid(ID, 11, 4)) 'תΪ1000��λ�����бȽ�
                    Dim i As Byte
                    For i = 1 To 12
                        If id4 >= xingzuo(i, 2) And id4 <= xingzuo(i, 3) Then
                            Info = xingzuo(i, 1)
                            Exit For
                        End If
                    Next
                Else
                    Info = errDataMsg
                End If
            Case 6    '6Ϊ�Ա�
                Info = IIf(Mid(ID, 17, 1) Mod 2, "��", "Ů")
            Case 7    '7�ж�У�����Ƿ���ȷ
                If CRC(ID) = UCase(Right(ID, 1)) Then
                    Info = True
                Else
                    Info = False
                End If
            Case 8  'У����
                Info = CRC(ID)
            Case 9  'ת18λ����
                Info = ToRule(ID)
            Case Else

        End Select
    Else

    End If

End Property

Property Get ToRule(ID As String)
'���ID���Ȳ�Ϊ15������ԭֵ
'���򷵻�18λ�ַ�

    If Len(ID) <> 15 Then
        ToRule = ID
    Else
        Dim pre17 As String
        pre17 = Left(ID, 6) & "19" & Right(ID, 9)
        ToRule = pre17 & CRC(pre17)
    End If

End Property

Property Get CRC(ID As String) As String
'��������֤У����
'���ID���Ȳ���17�򳬹�18������""
'����ֵ����Ϊ[0-9X]

    If Len(ID) = 17 Or Len(ID) = 18 Then
        Dim i As Byte, sum As Long
        For i = 1 To 17
            sum = sum + ((2 ^ (18 - i)) Mod 11) * Mid(ID, i, 1)
        Next
        Dim remainder As Long   '����
        remainder = sum Mod 11
        '0,1,2,3,4,5,6,7,8,9,10 => 1,0,X,9,8,7,6,5,4,3,2
        'm��0��1��2��3��4��5��6��7��8��9����10��10-m����
        'n��X��9��8��7��6��5��4��3��2��1��0(n + 2) mod 11 ��
        'r: 1��0��X��9��8��7��6��5��4��3��2
        If (12 - remainder) Mod 11 = 10 Then
            CRC = "X"
        Else
            CRC = (12 - remainder) Mod 11 & ""
        End If
    Else
        CRC = ""
    End If
End Property
