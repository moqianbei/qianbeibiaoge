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
    '提示用户
    If MsgBox("此操作可能会持续15秒左右，您确定要继续吗？", vbOKCancel) <> vbOK Then Exit Sub

    Dim t
    t = Timer

    Application.ScreenUpdating = False
    Excel.Application.DisplayAlerts = False '忽略弹窗警告

    '清除原来的所有数据
    premierSheet1.Cells.Clear:    premierSheet2.Cells.Clear:    premierSheet3.Cells.Clear
    If premierSheet0.FilterMode Then premierSheet0.Cells.AutoFilter '取消筛选状态，否则用end的方法无法得到所有数据，而是筛选过的数据
    '将所有区划sheet中 1级区划列 复制到一级区划sheet，并去除重复值
    premierSheet0.Range(premierSheet0.Range("c2"), premierSheet0.Range("c2").End(xlDown)).Copy premierSheet1.Range("A1")
    premierSheet1.Range(premierSheet1.Range("A1"), premierSheet1.Range("A1").End(xlDown)).removeDuplicates Columns:=1, Header:=xlNo
    premierSheet1.Range(premierSheet1.Range("A1"), premierSheet1.Range("A1").End(xlDown)).Copy
    '转置到二级区划表首行
    premierSheet2.[A1].PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, Transpose:=True

    Dim i As Long, col As Long, j As Long, k As Long
    col = premierSheet2.UsedRange.Columns.Count '一级区划数
    
    Dim temp As Worksheet
    Set temp = Sheets.Add(before:=Sheets(1))    '临时表，便于去除重复值
    '如果数据都在一个表中，VBA去重时会自动扩展选区，耗费资源导致程序运行崩溃
    
    For i = 1 To col

        premierSheet0.UsedRange.AutoFilter Field:=3, CRITERIA1:=Sheet7.Cells(1, i).Value

        temp.Columns(1).Delete  '清除数据，防止干扰后面数据
        premierSheet0.Range(premierSheet0.Range("D1").End(xlDown), premierSheet0.Range("D1").End(xlDown).End(xlDown)).Copy temp.[A1]
        temp.Range("A1:A" & temp.[A1].End(xlDown).Row).removeDuplicates Columns:=1, Header:=xlNo  '去除重复值
        temp.Range("A1:A" & temp.[A1].End(xlDown).Row).Copy premierSheet2.Cells(2, i)
        On Error Resume Next
        premierSheet2.Columns(i).SpecialCells(xlCellTypeBlanks).Delete  '删除空值单元格
    Next

    If premierSheet0.FilterMode Then premierSheet0.Cells.AutoFilter '清除筛选状态
    
    '复制数据到临时表，去重并清除空值所在行
    temp.Columns(Array(1, 2)).Delete    '清除之前的数据，防止干扰
    premierSheet0.Range("C1:D1", premierSheet0.Range("C1:D1").End(xlDown)).Copy temp.[A1]
    temp.Range("A1:B" & temp.[A1:B1].End(xlDown).Row).removeDuplicates Columns:=Array(1, 2), Header:=xlYes
    temp.Columns("B:B").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

    Dim allR As Variant, pre3 As Variant
    pre3 = temp.Range("A2:B" & temp.[A2:B2].End(xlDown).Row)
    allR = premierSheet0.UsedRange
    temp.Delete '删除临时表，数据均已存入数组

    For j = 1 To UBound(pre3)
        k = 2   '从第二行开始写入数据
        premierSheet3.Cells(1, j).Value = pre3(j, 1) & "_" & pre3(j, 2)    '首行写入“河北省_承德市”字样
        For i = 1 To UBound(allR)
            If allR(i, 3) = pre3(j, 1) And allR(i, 4) = pre3(j, 2) And allR(i, 5) <> "" Then
                premierSheet3.Cells(k, j).Value = allR(i, 5)
                k = k + 1
            End If
        Next
    Next

    Application.ScreenUpdating = True
    Excel.Application.DisplayAlerts = True
    MsgBox "已经重新生成，用时" & Timer - t & "秒"  '，自动为您更新公式名称"    13s

End Sub
