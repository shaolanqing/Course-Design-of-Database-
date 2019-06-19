VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame4 
      Caption         =   "修改删除"
      Height          =   2415
      Left            =   240
      TabIndex        =   19
      Top             =   7680
      Width           =   10695
      Begin VB.CommandButton Command7 
         Caption         =   "删除"
         Height          =   375
         Left            =   5520
         TabIndex        =   29
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "确定修改"
         Height          =   375
         Left            =   3240
         TabIndex        =   28
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "查询"
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   5760
         TabIndex        =   25
         Text            =   "Text9"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   3120
         TabIndex        =   23
         Text            =   "Text8"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   1080
         TabIndex        =   21
         Text            =   "Text7"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "总消费额"
         Height          =   180
         Left            =   4920
         TabIndex        =   24
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "会员类型"
         Height          =   180
         Left            =   2280
         TabIndex        =   22
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "卡号"
         Height          =   180
         Left            =   480
         TabIndex        =   20
         Top             =   480
         Width           =   360
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "添加"
      Height          =   2655
      Left            =   240
      TabIndex        =   10
      Top             =   4800
      Width           =   10815
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   5160
         TabIndex        =   37
         Text            =   "Text13"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   3360
         TabIndex        =   35
         Text            =   "Text12"
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1320
         TabIndex        =   33
         Text            =   "Text11"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   8760
         TabIndex        =   31
         Text            =   "Text10"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "检查数据再添加"
         Height          =   375
         Left            =   7800
         TabIndex        =   26
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   6480
         TabIndex        =   18
         Text            =   "Text6"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   4800
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3120
         TabIndex        =   14
         Text            =   "Text4"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "挂失"
         Height          =   180
         Left            =   4440
         TabIndex        =   36
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "充值"
         Height          =   180
         Left            =   2640
         TabIndex        =   34
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "管理员工号"
         Height          =   180
         Left            =   240
         TabIndex        =   32
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "总消费额"
         Height          =   180
         Left            =   7920
         TabIndex        =   30
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "余额"
         Height          =   180
         Left            =   6000
         TabIndex        =   17
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "折扣"
         Height          =   180
         Left            =   4320
         TabIndex        =   15
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "类型"
         Height          =   180
         Left            =   2280
         TabIndex        =   13
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "卡号"
         Height          =   180
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "查询"
      Height          =   4335
      Left            =   5880
      TabIndex        =   1
      Top             =   360
      Width           =   4695
      Begin VB.CommandButton Command3 
         Caption         =   "按会员类型查询"
         Height          =   495
         Left            =   3000
         TabIndex        =   9
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1440
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "按卡号查询"
         Height          =   495
         Left            =   2760
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1200
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "会员类型"
         Height          =   180
         Left            =   480
         TabIndex        =   7
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "卡号"
         Height          =   180
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "会员卡信息"
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      Begin VB.CommandButton Command1 
         Caption         =   "显示全部"
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   3600
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   240
         Top             =   3480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2175
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3836
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()              '按卡号查询
Dim sql As String
sql = "select * from card_number where cnumber='" + Text1.Text + "'"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = sql
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Command3_Click()              '按会员类型查询
Dim sql As String
sql = "select * from card_number where ctype='" + Text2.Text + "'"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = sql
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Command4_Click()                 '检查数据再添加
 Dim conn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   Dim sql As String
   connstr = "provider=SQLOLEDB.1;User ID=toto;pwd=123;initial catalog=mem;Data source=DESKTOP-K8KU961"
   sql = " select * from card_number where cnumber='" + Text3.Text + "'"
   conn.Open connstr
   rs.Open sql, conn, 3, 2
   If Not rs.EOF Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Then
     MsgBox ("学号已存在或数据项为空!!!")
     Text3.Text = ""
     Text4.Text = ""
     Text5.Text = ""
     Text6.Text = ""
     Text10.Text = ""
     Text11.Text = ""
     Text12.Text = ""
     Text13.Text = ""
   Else
     sql = "insert into card_number values('" + Trim(Text3.Text) + "','" + Trim(Text4.Text) + "','" + Trim(Text5.Text) + "','" + Trim(Text6.Text) + "','" + Trim(Text10.Text) + "','" + Trim(Text11.Text) + "', '" + Trim(Text12.Text) + "','" + Trim(Text13.Text) + "')"
     MsgBox (sql)
     rs.Close
     
     rs.Open sql, conn, 1, 1
      MsgBox ("已成功添加")
     Adodc1.Refresh
   Set DataGrid1.DataSource = Adodc1
     
   End If
End Sub

Private Sub Command5_Click()                     '查询
  Dim conn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
     Dim sql As String
     connstr = "provider=SQLOLEDB.1;User ID=toto;pwd=123;initial catalog=mem;Data source=DESKTOP-K8KU961"
   sql = "select * from card_number where cnumber='" + Text7.Text + "'"
   conn.Open connstr
   rs.Open sql, conn, 3, 3
   If Not rs.EOF Then
   Text7.Text = rs("cnumber")
   Text8.Text = rs("ctype")
   Text9.Text = rs("all_consume")
   Text9.Enabled = False
   flag = 1
   Else
    MsgBox ("查无此人")
    flag = 0
   End If
End Sub

Private Sub Command6_Click()                       '确定修改
 Dim conn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
     Dim sql As String
     If flag = 2 Then
     MsgBox ("先查找学生是否存在！")
      ElseIf flag = 0 Then
     MsgBox ("该生不存在")
     Else
   sql = "update card_number set cnumber='" + Text7.Text + "',ctype='" + Text8.Text + "',all_consume='" + Text9.Text + "' where cnumber='" + Text7.Text + "'"
   conn.Open connstr
   rs.Open sql, conn, 3, 3
   MsgBox ("已成功更新")
   Adodc1.Refresh
   Set DataGrid1.DataSource = Adodc1
  End If
End Sub

Private Sub Command7_Click()                        '删除
Dim conn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   Dim sql As String
   If flag = 2 Then
     MsgBox ("先查询该生是否存在")
    ElseIf flag = 0 Then
     MsgBox ("该生不存在")
     Else
     
   sql = " delete from card_number where cnumber='" + Text7.Text + "'"
   conn.Open connstr
   rs.Open sql, conn, 3, 2
    MsgBox ("已成功删除")
   Adodc1.Refresh
   Set DataGrid1.DataSource = Adodc1
   End If
   Form1.Show
End Sub


