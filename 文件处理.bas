Dim auto_workname, auto_filepath, auto_filename, auto_houzhui, auto_shijian
Public newtime


Sub 自动备份当前文件()
'目前仅支持一个文件

'获取当前文件后缀名，防止文件名中有.
auto_workname = ActiveWorkbook.Name

arr = Split(auto_workname, ".")
auto_houzhui = arr(UBound(arr))
auto_filepath = strFolder("请选择要自动保存到哪个文件夹")
auto_filename = Split(auto_workname, ".")(0)
auto_shijian = InputBox("请设置自动保存的时间间隔(格式：00:00:05)")
 
Application.OnTime Now + TimeValue(auto_shijian), "调用自动保存"

End Sub

Sub autoSave(workname, filepath, fileName, houzhui)

Excel.Application.DisplayAlerts = False '避免弹出个人信息确认框
Workbooks(workname).SaveCopyAs fileName:=filepath & fileName & "-" & Format(Now(), "yyyymmddhhmmss") & "." & houzhui
Excel.Application.DisplayAlerts = True
End Sub
Sub 调用自动保存()

Call autoSave(auto_workname, auto_filepath, auto_filename, auto_houzhui)

newtime = Now + TimeValue(auto_shijian)
Application.OnTime newtime, "调用自动保存"
End Sub

Sub 取消定时保存()

On Error GoTo 100
'Application.OnTime Now + TimeValue("00:00:05"), "调用自动保存", , False
Application.OnTime newtime, "调用自动保存", , False
Exit Sub
100:
MsgBox "请先按开启定时保存!"

End Sub

Sub 文件批量重命名()
Dim str

flag = 0
For Each sht In ActiveWorkbook.Sheets
    If sht.Name = "文件批量重命名" Then
        flag = 1
    Exit For
    End If
Next

If flag = 0 Then    '如果不存在，则复制
    ThisWorkbook.Sheets("文件批量重命名").Copy before:=ActiveWorkbook.Sheets(1)
    MsgBox "表格模板已复制，下面请选择需要重命名的文件"
End If

With ActiveWorkbook.Sheets("文件批量重命名")
    If .Range("a3") = "" Then

        On Error Resume Next '加上这句防止用户点击取消发生的错误
        str = Application.GetOpenFilename("所有文件,*.*", Title:="请选择要重命名的文件", MultiSelect:=True)
        Application.ScreenUpdating = False

            For i = LBound(str) To UBound(str)
                .Range("a" & i + 2) = str(i)
            Next
        Application.ScreenUpdating = True

        MsgBox "现在请打开""文件批量重命名""表格进行编辑，编辑完成后再次点击此程序即可"
        .Select
        
    Else
        '重命名操作
        
        t = Timer
        
        For i = 3 To 65536
        
        If Range("a" & i) <> "" Then
            On Error Resume Next '为防止重名文件
            Name Range("a" & i) As Range("h" & i)
        Else
            Exit For
        End If
        Next
        
        '删除多余数据
                
            .ListObjects("表1").Resize Range("$A$2:$H$3")

            .Range("表1[完整路径]").ClearContents
            .Range("表1[路径2]").ClearContents
            .Range("表1[文件名3]").ClearContents
            .Range("表1[文件扩展名]").ClearContents

            .Range("A4").Select
            Range(Selection, Selection.End(xlToRight)).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.ClearContents
            
            .Range("a3").Select

        MsgBox "已完成重命名操作，共用时" & Timer - t & "秒！"
        
    End If
End With

End Sub


