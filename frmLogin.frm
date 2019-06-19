VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "登录"
   ClientHeight    =   8025
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   4741.436
   ScaleMode       =   0  'User
   ScaleWidth      =   12985.62
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   3960
      TabIndex        =   1
      Top             =   1680
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   390
      Left            =   3120
      TabIndex        =   4
      Top             =   3120
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   390
      Left            =   5040
      TabIndex        =   5
      Top             =   3120
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "用户名称(&U):"
      Height          =   270
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "密码(&P):"
      Height          =   270
      Index           =   1
      Left            =   3000
      TabIndex        =   2
      Top             =   2400
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    '设置全局变量为 false
    '不提示失败的登录
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    '检查正确的密码
    If txtPassword = "123" Then
        '将代码放在这里传递
        '成功到 calling 函数
        '设置全局变量时最容易的
        LoginSucceeded = True
        Me.Hide
        Form3.Show
    Else
        MsgBox "无效的密码，请重试!", , "登录"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub
