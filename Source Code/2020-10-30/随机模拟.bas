

'统计某字符的数目
Function CountX(strInput As String, strFind As String) As Integer
    CountX = Len(strInput) - Len(Replace(strInput, strFind, ""))
End Function
'1为男，0为女
Function J_SJXM(Optional xingbie As Byte = 2)

Dim xingming As String

If xingbie = 1 Then
    xingming = Split(xingshi, ",")(Int(Rnd * CountX(xingshi, ",") + 1)) & _
    Split(nanxingming, ",")(Int(Rnd * CountX(nanxingming, ",") + 1))

ElseIf xingbie = 0 Then
    xingming = Split(xingshi, ",")(Int(Rnd * CountX(xingshi, ",") + 1)) & _
    Split(nvxingming, ",")(Int(Rnd * CountX(nvxingming, ",") + 1))

Else
    xingming = Split(xingshi, ",")(Int(Rnd * CountX(xingshi, ",") + 1)) & _
    Split(nanxingming & nvxingming, ",")(Int(Rnd * CountX(nanxingming & nvxingming, ",") + 1))
End If
J_SJXM = xingming
End Function


