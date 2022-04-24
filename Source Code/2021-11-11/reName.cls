VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "文件批量重命名"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButtonReChoose_Click()

With ActiveWorkbook.Sheets("文件批量重命名")
    Dim str
    On Error Resume Next '加上这句防止用户点击取消发生的错误
    str = Application.GetOpenFilename("所有文件,*.*,PDF文件,*.pdf;*.PDF,Word文档,*.doc*,Excel表格,*.xls*", FilterIndex:=1, Title:="请选择要重命名的文件", MultiSelect:=True)
    If Err Then MsgBox "已取消操作，程序已中止": Exit Sub
    Application.ScreenUpdating = False

    .Range(.Range("A4:h4"), .Range("A4:h4").End(xlDown)).Clear
    .Range("a3") = "M:\浅北表格助手\示例文档.xlsm"

    For i = LBound(str) To UBound(str)
        .Range("a" & i + 3) = str(i)    '从第4行单元格开始读取选择的文件
    Next
    '自动填充公式
    .Range("B3:H3").AutoFill destination:=.Range("B3:H" & UBound(str) + 3), Type:=xlFunction

    Application.ScreenUpdating = True

    MsgBox "编辑完成后点击上方开始按钮即可"
    .Select
End With
      
End Sub

Private Sub CommandButtonRename_Click()
' 重命名操作

    If Range("a3") = "M:\浅北表格助手\示例文档.xlsm" Then
        MsgBox "未选择任何文件，程序已退出"
        Exit Sub
    Else
        t = Timer
        With ActiveWorkbook.Sheets("文件批量重命名")
            For i = 3 To Range("a3").End(xlDown).Row
                On Error Resume Next '为防止重名文件
                If Range("A" & i).Value = "" Then Exit For
                Name Range("a" & i) As Range("h" & i)
            Next
        End With
        MsgBox "已完成重命名操作，共用时" & Timer - t & "秒！"
    End If
End Sub

