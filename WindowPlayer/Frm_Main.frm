VERSION 5.00
Begin VB.Form Frm_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "玩窗体"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   6270
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame F2 
      Caption         =   "标题"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   6015
      Begin VB.CommandButton Cmd_ChangeTlt 
         Caption         =   "更改"
         Height          =   300
         Left            =   5160
         TabIndex        =   7
         Top             =   337
         Width           =   735
      End
      Begin VB.TextBox Txt_Tlt 
         Height          =   270
         Left            =   960
         TabIndex        =   6
         Top             =   352
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "新标题："
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame F1 
      Caption         =   "获取窗体"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton Cmd_Catch 
         Caption         =   "单击后去到目标窗口"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   5535
      End
      Begin VB.TextBox Txt_FintTlt 
         Height          =   270
         Left            =   1440
         TabIndex        =   3
         Top             =   232
         Width           =   3735
      End
      Begin VB.CommandButton Cmd_Search 
         Caption         =   "查找"
         Height          =   255
         Left            =   5160
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Lbl_Title 
         Caption         =   "当前窗口"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "输入窗体标题："
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private MyHwnd As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long



Private Sub Cmd_Catch_Click()
    Dim i As Integer
    For i = 0 To 4
        Cmd_Catch.Caption = "你有" & (5 - i) & "秒到达目标窗口"
        DoEvents
        Sleep 1000
    Next
    MyHwnd = GetForegroundWindow
    If MyHwnd <> 0 Then
        Dim temp As String
        temp = String(255, 0)
        GetWindowText MyHwnd, temp, 255
        MsgBox "查找窗体成功", vbOKOnly, "成功"
        Lbl_Title.Caption = "当前窗口:" + temp
    Else
        MsgBox "查找窗体失败", vbOKOnly, "失败"
    End If
    Cmd_Catch.Caption = "单击后去到目标窗口"
End Sub

Private Sub Cmd_ChangeTlt_Click()
    If MyHwnd <> 0 Then
        SetWindowText MyHwnd, Txt_Tlt.Text
    End If
End Sub

Private Sub Cmd_Search_Click()
    MyHwnd = FindWindow(vbNullString, Txt_FintTlt)
    If MyHwnd <> 0 Then
         Dim temp As String
        temp = String(255, 0)
        GetWindowText MyHwnd, temp, 255
        MsgBox "查找窗体成功", vbOKOnly, "成功"
        Lbl_Title.Caption = "当前窗口:" + temp
    Else
        MsgBox "查找窗体失败，确定标题没写错吗?", vbOKOnly, "失败"
    End If
End Sub
