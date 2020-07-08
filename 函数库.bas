Function JVLOOKUP(str As String, Ran As Range, i As Integer)
'使用vba的find函数，实现工作表的match、index功能

Dim rng As Range

Set rng = Ran.Find(str, , , xlWhole)

If Not rng Is Nothing Then

    JVLOOKUP = rng.Offset(0, i).value '右侧为正，左侧为负

End If

'未找到会返回数字0
End Function
'中国式排名12234445
Function JPAIMING(igji As Double, quyu As Range, Optional order As Byte = 1)

    On Error Resume Next
    Dim rng, i As Integer, Only As New Collection

    For Each rng In quyu
        '如果rng 大于成绩，则将其导入到 only集合中，用于去除重复值
        If order = 0 Then
            If rng > igji Then Only.Add rng, CStr(rng)
        Else
            If rng < igji Then Only.Add rng, CStr(rng)
        End If
    Next

    '将集合only的数据个数加1后赋值给函数，作为做终结果

    JPAIMING = Only.Count + 1

End Function
'条件排名
'英式排名=SUMPRODUCT((条件区域1=条件1)* (条件区域2=条件2)* (数据区域>数据))
Function PM(igji As Double, quyu As Range, Optional criferia1_range As Range, Optional criferia1 As Range, Optional criferia2_range As Range = 1, Optional criferia2 As Range, Optional criferia3_range As Range = 1, Optional criferia3 As Range = 1)
    'Set zhipianyi = quyu.Offset(1, 0).Resize(quyu.Rows.Count + 1)
    PM.FormulaArray = " Sum((Frequency((criferia1_range = criferia1) * (criferia1_range = criferia2) * (criferia3_range = criferia3) * 值所在区域, 值所在区域) > 0) * (qy.Offset(-1, 0).Resize(qy.Rows.Count + 1) > igji)) + 1"
End Function
Function JHYPELINK(rng)
'获取某单元格的内部链接
    Application.Volatile True
    With rng.Hyperlinks(1)
        JHYPELINK = IIf(.Address = "", .SubAddress, .Address)
    End With
End Function


'提取身份证信息
'1为地区
'2为出生日期
'3为年龄
'4为生肖
'5为星座
'6为性别
'7判断校验码是否正确

Function JSHENFENZHENG(ID As String, getType As Integer)

Dim info As Variant

If Len(ID) <> 18 Then
    info = "身份证位数不正确"
Else

    Dim dizhima As String, id6 As String, id4 As Variant
    Dim shengxiao(), xingzuo
    shengxiao = Array("鼠", "牛", "虎", "兔", "龙", "蛇", "马", "羊", "猴", "鸡", "狗", "猪")
    xingzuo = [{"水瓶座",0121,0219;"双鱼座",0220,0320;"白羊座",0321,0420;"金牛座",0421,0521;"双子座",0522,0621;"巨蟹座",0622,0723;"狮子座",0724,0823;"处女座",0824,0923;"天秤座",0924,1023;"天蝎座",1024,1122;"射手座",1123,1222;"摩羯座",1223,0122}]

    Select Case getType
        Case 1
            dizhima = DATA_QUYUMA
            id6 = Left(ID, 6)
            
            '防止出现找不到的错误
            If VBA.Strings.InStr(dizhima, id6) <> 0 Then
                dizhima = Split(dizhima, id6)(1)
                info = Split(dizhima, ",")(1)
            Else
                info = "未找到籍贯"
            End If

        Case 2
            If Format(DateSerial(Mid(ID, 7, 4), Mid(ID, 11, 2), Mid(ID, 13, 2)), "yyyymmdd") = Mid(ID, 7, 8) Then
                info = DateSerial(Mid(ID, 7, 4), Mid(ID, 11, 2), Mid(ID, 13, 2))
            Else
                info = "身份证出生日期有误"
            End If

        Case 3
            info = DateDiff("yyyy", JSHENFENZHENG(ID, 2), Date)

        Case 4
            info = shengxiao((Mid(ID, 7, 4) - 1900) Mod 12)

        Case 5
        id4 = Val(Mid(ID, 11, 4)) '转为1000四位数进行比较
            For i = 1 To 12
                If id4 >= xingzuo(i, 2) And id4 <= xingzuo(i, 3) Then
                    info = xingzuo(i, 1)
                    Exit For
                End If
            Next

        Case 6
            info = IIf(Mid(ID, 17, 1) Mod 2 = 0, "女", "男")
        
        Case 7
            Dim sump As Integer
            sump = 0
            For i = 18 To 1 Step -1
                Dim Wi, Ai
                If Mid(ID, i, 1) = "x" Or Mid(ID, i, 1) = "X" Then
                    Ai = 10
                Else
                    Ai = Mid(ID, i, 1)
                End If
                he = Ai * (WorksheetFunction.Power(2, 18 - i) Mod 11)
                sump = sump + he
            Next
            
            If sump Mod 11 = 1 Then
                info = "检验码通过验证"
            Else
                info = "校验码出错"
            End If
        End Select
End If
JSHENFENZHENG = info
End Function
'自动获取校验码，方便身份证号模拟
Function JJYM(IDs As Range, Optional IDs2 As String, Optional IDs3 As String, Optional IDs4 As String, Optional IDs5 As String, Optional IDs6 As String)

Dim ID17 As String
Dim rng As Range
ID17 = ""

If IDs2 = "" Then
    For Each rng In IDs
        ID17 = ID17 & rng.value
    Next
Else
    ID17 = IDs & IDs2 & IDs3 & IDs4 & IDs5 & IDs6
End If

If Len(ID17) <> 17 Then
    JJYM = "请重新选择"
Else
    
    flag = "X" '标记检验码是否为数字，默认为X
    For i = 0 To 9
        If JSHENFENZHENG(ID17 & Format(i, "@"), 7) = "检验码通过验证" Then
            flag = Format(i, "@")
            Exit For
        End If
    Next
    JJYM = flag
End If
End Function
