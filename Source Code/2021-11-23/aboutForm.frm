VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} aboutForm 
   Caption         =   "浅北表格 - 关于"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9675.001
   OleObjectBlob   =   "aboutForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "aboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LabelGitee_Click()
    ActiveWorkbook.FollowHyperlink Address:="https://gitee.com/mo-qianbei/qianbeibiaoge", NewWindow:=True
End Sub

Private Sub LabelYueQue_Click()
    ActiveWorkbook.FollowHyperlink Address:="https://www.yuque.com/moqianbei/qianbeibiaoge", NewWindow:=True
End Sub

Private Sub UserForm_Activate()
    Call loadFun
End Sub
