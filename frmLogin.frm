VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��¼"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   3960
      TabIndex        =   1
      Top             =   1680
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   390
      Left            =   3120
      TabIndex        =   4
      Top             =   3120
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
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
      Caption         =   "�û�����(&U):"
      Height          =   270
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "����(&P):"
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
    '����ȫ�ֱ���Ϊ false
    '����ʾʧ�ܵĵ�¼
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    '�����ȷ������
    If txtPassword = "123" Then
        '������������ﴫ��
        '�ɹ��� calling ����
        '����ȫ�ֱ���ʱ�����׵�
        LoginSucceeded = True
        Me.Hide
        Form3.Show
    Else
        MsgBox "��Ч�����룬������!", , "��¼"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub
