Attribute VB_Name = "Fun"
Option Explicit '强制变量声明
'Option Base 0 '数组下标从0开始

Function JVLOOKUP(Lookup_value As String, Table_array As Range, Offset_num As Integer)
Attribute JVLOOKUP.VB_Description = "搜索表区域中满足条件的第一个单元格。如果找到，返回它右侧第±N个单元格的值；未找到返回0"
Attribute JVLOOKUP.VB_ProcData.VB_Invoke_Func = " \n20"

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
Attribute JRANK.VB_Description = "【中国式排名】返回某数字在一列数字中相对于其他数值的大小排名。排名并列相同，不占用名次。\neg：有并列第2，则第4名自动变为第3名（122345）"
Attribute JRANK.VB_ProcData.VB_Invoke_Func = " \n20"
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
Attribute JHYPELINK.VB_Description = "【提取链接】返回单元格设置的链接，如果未找到，返回错误值"
Attribute JHYPELINK.VB_ProcData.VB_Invoke_Func = " \n20"
'获取某单元格的内部链接
    Application.Volatile True
    With Ref.Hyperlinks(1)
        JHYPELINK = IIf(.Address = "", .SubAddress, .Address)
    End With
End Function
Function JSHENFENZHENG(ID As String, Optional getType As Integer = 7)
Attribute JSHENFENZHENG.VB_Description = "【身份证信息】获取中国大陆居民身份证包含的信息"
Attribute JSHENFENZHENG.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim shenfenID As New IDCard
    JSHENFENZHENG = shenfenID.Info(ID, getType)
End Function

'1为男，0为女，2随机
Function JRANKNAME(Optional ByVal Sex As Byte = 2)
Attribute JRANKNAME.VB_Description = "【随机姓名】返回一个随机的中文姓名"
Attribute JRANKNAME.VB_ProcData.VB_Invoke_Func = " \n20"

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
    JRANKNAME = xingming
End Function

