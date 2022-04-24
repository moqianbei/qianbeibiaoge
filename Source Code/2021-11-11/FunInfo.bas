Attribute VB_Name = "函数介绍"
Option Explicit
'函数功能介绍，不能放在程序启动时加载，需要点击按钮调用
Sub loadFun()

    Dim JVLOOKUPArg(1 To 3) As String
    JVLOOKUPArg(1) = "【查找值】需要在数据表进行搜索的值，可以是数值、引用或字符串"
    JVLOOKUPArg(2) = "【搜索区域】在哪里找"
    JVLOOKUPArg(3) = "【相对位置】如果找到了该数据，要返回它右侧的第几个单元格的数据。0为数据本身，负数表示向左查找"
    Excel.Application.MacroOptions "JVLOOKUP", _
    "搜索表区域中满足条件的第一个单元格。如果找到，返回它右侧第±N个单元格的值；未找到返回0", _
    Category:="浅北表格助手", _
    ArgumentDescriptions:=JVLOOKUPArg

'-------------------------------------------------------------------------------------------------

    Dim JRANKArg(1 To 3) As String
    JRANKArg(1) = "【数字】找哪个数字的排名"
    JRANKArg(2) = "【区域】这个数在哪个区域的排名"
    JRANKArg(3) = "[排序方式] 指定排名的方式。0或忽略，降序（值越大排名越靠前）；1:反之"
    Excel.Application.MacroOptions "JRANK", _
    "【中国式排名】返回某数字在一列数字中相对于其他数值的大小排名。排名并列相同，不占用名次。" & Chr(10) & "eg：有并列第2，则第4名自动变为第3名（122345）", _
    Category:="浅北表格助手", _
    ArgumentDescriptions:=JRANKArg

'-------------------------------------------------------------------------------------------------

    Dim JSHENFENZHENGArg(1 To 2) As String
    JSHENFENZHENGArg(1) = "【身份证号】15位或18位字符"
    JSHENFENZHENGArg(2) = "[2信息类型] 获取身份证中的哪种信息" & Chr(10) & _
      "1:地区  2:生日  3:年龄  4:生肖  5:星座  6:性别" & Chr(10) & _
      "7:是否合规（默认）  8:校验码  9:转18位号码"
    Excel.Application.MacroOptions "JSHENFENZHENG", _
    "【身份证信息】获取中国大陆居民身份证包含的信息", _
    Category:="浅北表格助手", _
    ArgumentDescriptions:=JSHENFENZHENGArg

'-------------------------------------------------------------------------------------------------

    Excel.Application.MacroOptions "JHYPELINK", _
    "【提取链接】返回单元格设置的链接，如果未找到，返回错误值", _
    Category:="浅北表格助手", _
    ArgumentDescriptions:="需要获取链接的单元格"

'-------------------------------------------------------------------------------------------------

    Excel.Application.MacroOptions "JRANKNAME", _
    "【随机姓名】返回一个随机的中文姓名", _
    Category:="浅北表格助手", _
    ArgumentDescriptions:="[性别] 1:男  0:女  2:随机（默认值）"

End Sub
