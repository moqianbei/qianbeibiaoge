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
'提取身份证信息
Property Get Info(ID As String, Optional getType As Integer = 7) As Variant
'ID不为15位或18位，返回空
'1为地区，未找到返回空
'2为真正出生日期，如果该数字错误返回，2及以下值均返回错误提示字符串
    '3为年龄（周岁数字）
    '4为生肖
    '5为星座
'6为性别：男女
'7判断校验码是否正确：布尔值Boolean

    Info = "身份证号不符合规范"   '设定默认返回值
    
    If Len(ID) = 15 Then
        Info = Info(ToRule(ID), getType)
    ElseIf Len(ID) = 18 Then

        Dim errDataMsg As String, errPlaceMsg As String
        errDataMsg = "身份证出生日期有误"
        errPlaceMsg = "找不到身份证所在地区"

        Select Case getType
            Case 1    '1为地区
                Dim pre6 As String
                pre6 = Left(ID, 6)
                On Error Resume Next
                '单元格匹配
                Info = premierSheet0.Range("A:A").Find(What:=pre6, LookAt:=xlWhole).Offset(0, 1).Value
                If Err.Number = 91 Then
                    Info = errPlaceMsg
                    Err.Clear
                End If
                On Error GoTo 0

            Case 2  '出生日期转换为真正日期
                If Format(DateSerial(Mid(ID, 7, 4), Mid(ID, 11, 2), Mid(ID, 13, 2)), "yyyymmdd") = Mid(ID, 7, 8) Then
                    Info = DateSerial(Mid(ID, 7, 4), Mid(ID, 11, 2), Mid(ID, 13, 2))
                Else
                    Info = errDataMsg
                End If
    
            Case 3    '3为年龄
                If Info(ID, 2) <> errDataMsg Then
                    Info = DateDiff("yyyy", Info(ID, 2), Date)
                Else
                    Info = errDataMsg
                End If
            Case 4    '4为生肖
                Dim shengxiao() As Variant
                shengxiao = Array("鼠", "牛", "虎", "兔", "龙", "蛇", "马", "羊", "猴", "鸡", "狗", "猪")

                If Info(ID, 2) <> errDataMsg Then
                    Info = shengxiao((Mid(ID, 7, 4) - 1900) Mod 12) '1900年为鼠年
                Else
                    Info = errDataMsg
                End If
            Case 5    '5为星座
                If Info(ID, 2) <> errDataMsg Then
                    Dim xingzuo
                    xingzuo = [{"水瓶座",0121,0219;"双鱼座",0220,0320;"白羊座",0321,0420;"金牛座",0421,0521;"双子座",0522,0621;"巨蟹座",0622,0723;"狮子座",0724,0823;"处女座",0824,0923;"天秤座",0924,1023;"天蝎座",1024,1122;"射手座",1123,1222;"摩羯座",1223,0122}]
                    Dim id4
                    id4 = Val(Mid(ID, 11, 4)) '转为1000四位数进行比较
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
            Case 6    '6为性别
                Info = IIf(Mid(ID, 17, 1) Mod 2, "男", "女")
            Case 7    '7判断校验码是否正确
                If CRC(ID) = UCase(Right(ID, 1)) Then
                    Info = True
                Else
                    Info = False
                End If
            Case 8  '校验码
                Info = CRC(ID)
            Case 9  '转18位号码
                Info = ToRule(ID)
            Case Else

        End Select
    Else

    End If

End Property

Property Get ToRule(ID As String)
'如果ID长度不为15，返回原值
'否则返回18位字符

    If Len(ID) <> 15 Then
        ToRule = ID
    Else
        Dim pre17 As String
        pre17 = Left(ID, 6) & "19" & Right(ID, 9)
        ToRule = pre17 & CRC(pre17)
    End If

End Property

Property Get CRC(ID As String) As String
'返回身份证校验码
'如果ID长度不足17或超过18，返回""
'返回值可能为[0-9X]

    If Len(ID) = 17 Or Len(ID) = 18 Then
        Dim i As Byte, sum As Long
        For i = 1 To 17
            sum = sum + ((2 ^ (18 - i)) Mod 11) * Mid(ID, i, 1)
        Next
        Dim remainder As Long   '余数
        remainder = sum Mod 11
        '0,1,2,3,4,5,6,7,8,9,10 => 1,0,X,9,8,7,6,5,4,3,2
        'm：0－1－2－3－4－5－6－7－8－9——10（10-m）↓
        'n：X－9－8－7－6－5－4－3－2－1－0(n + 2) mod 11 ↓
        'r: 1－0－X－9－8－7－6－5－4－3－2
        If (12 - remainder) Mod 11 = 10 Then
            CRC = "X"
        Else
            CRC = (12 - remainder) Mod 11 & ""
        End If
    Else
        CRC = ""
    End If
End Property

