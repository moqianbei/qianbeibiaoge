Attribute VB_Name = "Fun"
Option Explicit '强制变量声明
'Option Base 0 '数组下标从0开始

Function JVLOOKUP(Lookup_value As String, Table_array As Range, Offset_num As Integer)

    '使用vba的find函数，实现工作表的match、index功能
    '未找到会返回数字0
    
    Dim rng As Range
    Set rng = Table_array.Find(Lookup_value, , , xlWhole)
    If Not rng Is Nothing Then
        JVLOOKUP = rng.Offset(0, Offset_num).Value '右侧为正，左侧为负
    End If

End Function
'中国式排名12234445
Function JRANK(Number As Double, Ref As Range, Optional Order = 0)
    On Error Resume Next
    Dim rng, i As Integer, Only As New Collection
    For Each rng In Ref
        '如果rng 大于成绩，则将其导入到 only 集合中，用于去除重复值
        If Order = False Then   '为 0 或忽略，降序，数字越大，排名越好
            If rng > Number Then Only.Add rng, CStr(rng)
        Else            '升序，数字越大，排名越靠后
            If rng < Number Then Only.Add rng, CStr(rng)
        End If
    Next
    '将集合only的数据个数加1后赋值给函数，作为做终结果
    JRANK = Only.Count + 1
End Function
'条件排名
'英式排名=SUMPRODUCT((条件区域1=条件1)* (条件区域2=条件2)* (数据区域>数据))
Function PM(igji As Double, quyu As Range, Optional criferia1_range As Range, Optional criferia1 As Range, Optional criferia2_range As Range = 1, Optional criferia2 As Range, Optional criferia3_range As Range = 1, Optional criferia3 As Range = 1)
    'Set zhipianyi = quyu.Offset(1, 0).Resize(quyu.Rows.Count + 1)
    PM.FormulaArray = " Sum((Frequency((criferia1_range = criferia1) * (criferia1_range = criferia2) * (criferia3_range = criferia3) * 值所在区域, 值所在区域) > 0) * (qy.Offset(-1, 0).Resize(qy.Rows.Count + 1) > igji)) + 1"
End Function
Function JHYPELINK(Ref)
'获取某单元格的内部链接
    Application.Volatile True
    With Ref.Hyperlinks(1)
        JHYPELINK = IIf(.Address = "", .SubAddress, .Address)
    End With
End Function
Function JSHENFENZHENG(ID As String, Optional getType As Integer = 7)
    Dim shenfenID As New IDCard
    JSHENFENZHENG = shenfenID.Info(ID, getType)
End Function

'1为男，0为女，2随机
Function JRANDNAME(Optional ByVal Sex As Byte = 2)

    Dim xingming As String, xingshi As String, nanxingming As String, nvxingming As String
    
    '避免重复调用影响运行时间，先进行读取
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
    JRANDNAME = xingming
End Function

'截取文本中两个分隔符之间的文本
Function JSPLIT(find_text As Range, splitter1 As String, splitter2 As String) As String

    Dim temp As String
    '如果可以找到分隔的字符串1，那么取后半段
    If InStr(find_text, splitter1) <> 0 Then
        temp = Split(find_text, splitter1)(1)
        If InStr(temp, splitter2) <> 0 Then
            JSPLIT = Split(temp, splitter2)(0)
        Else
            JSPLIT = "第2分隔符不正确"
        End If
    Else
        JSPLIT = "第1分隔符不正确"
    End If

End Function

'替换文本内容
Function JREPLACE(old_text As Range, find_text As String, replace_text As String) As String

    JREPLACE = Replace(old_text, find_text, replace_text)

End Function



