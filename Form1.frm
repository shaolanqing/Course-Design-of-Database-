VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   14310
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame4 
      Caption         =   "修改删除"
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Top             =   6360
      Width           =   12495
      Begin VB.CommandButton Command6 
         Caption         =   "删除"
         Height          =   495
         Left            =   6120
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "确定修改"
         Height          =   420
         Left            =   4320
         TabIndex        =   27
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "查询"
         Height          =   420
         Left            =   2280
         TabIndex        =   26
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   390
         Left            =   6240
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   3360
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   960
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   180
         Left            =   5280
         TabIndex        =   24
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   180
         Left            =   2760
         TabIndex        =   22
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "身份证号："
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "添加"
      Height          =   1455
      Left            =   0
      TabIndex        =   9
      Top             =   4920
      Width           =   12375
      Begin VB.TextBox Text11 
         Height          =   390
         Left            =   5280
         TabIndex        =   33
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   3360
         TabIndex        =   31
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "检查数据再添加"
         Height          =   495
         Left            =   7680
         TabIndex        =   18
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   390
         Left            =   1200
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   390
         Left            =   5160
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   390
         Left            =   3240
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1200
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "性别："
         Height          =   180
         Left            =   4680
         TabIndex        =   32
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "消费金额："
         Height          =   180
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "消费日期："
         Height          =   180
         Left            =   2520
         TabIndex        =   16
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "卡号："
         Height          =   180
         Left            =   4680
         TabIndex        =   14
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "姓名："
         Height          =   180
         Left            =   2640
         TabIndex        =   12
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "身份证号："
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "查询"
      Height          =   4695
      Left            =   7920
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command7 
         Caption         =   "按姓名查询"
         Height          =   420
         Left            =   2880
         TabIndex        =   29
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1200
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "按卡号查询"
         Height          =   495
         Left            =   2760
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   180
         Left            =   480
         TabIndex        =   7
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "卡号"
         Height          =   180
         Left            =   480
         TabIndex        =   4
         Top             =   960
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "会员信息显示"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton Command1 
         Caption         =   "显示全部"
         Height          =   495
         Left            =   4440
         TabIndex        =   2
         Top             =   3960
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   390
         Left            =   360
         Top             =   3960
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   688
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
         Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=mem;Data Source=127.0.0.1"
         OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=mem;Data Source=127.0.0.1"
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
         Height          =   3135
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5530
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public flag As Integer
Private Sub Command1_Click()                      '显示全部信息
Adodc1.CommandType = adCmdTable
Adodc1.RecordSource = "member"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Command2_Click()               '按卡号查询
Dim sql As String
sql = "select * from card_number where cnumber='" + Text1.Text + "'"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = sql
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub Command3_Click()                 '检查数据再添加
Dim conn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   Dim sql As String
   connstr = "provider=SQLOLEDB.1;User ID=toto;pwd=123;initial catalog=mem;Data source=127.0.0.1"
   sql = " select * from member where id='" + Text3.Text + "'"
   conn.Open connstr
   rs.Open sql, conn, 3, 2
   If Not rs.EOF Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text10.Text = "" Or Text11.Text = "" Then
     MsgBox ("学号已存在或数据项为空!!!")
     Text3.Text = ""
     Text4.Text = ""
     Text5.Text = ""
     Text6.Text = ""
     Text10.Text = ""
     Text11.Text = ""
   Else
     sql = "insert into member values('" + Trim(Text3.Text) + "','" + Trim(Text4.Text) + "','" + Trim(Text5.Text) + "','" + Trim(Text6.Text) + "', '" + Trim(Text10.Text) + "','" + Trim(Text11.Text) + "')"
     MsgBox (sql)
     rs.Close
     
     rs.Open sql, conn, 1, 1
      MsgBox ("已成功添加")
     Adodc1.Refresh
   Set DataGrid1.DataSource = Adodc1
     
   End If
End Sub

Private Sub Command4_Click()                      '查询
 Dim conn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
     Dim sql As String
     connstr = "provider=SQLOLEDB.1;User ID=toto;pwd=123;initial catalog=mem;Data source=DESKTOP-K8KU961"
   sql = "select * from member where id='" + Text7.Text + "'"
   conn.Open connstr
   rs.Open sql, conn, 3, 3
   If Not rs.EOF Then
   Text7.Text = rs("id")
   Text8.Text = rs("name")
   Text9.Text = rs("sex")
   Text9.Enabled = False
   flag = 1
   Else
    MsgBox ("查无此人")
    flag = 0
   End If
End Sub

Private Sub Command5_Click()                            '确定修改
Dim conn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
     Dim sql As String
     If flag = 2 Then
     MsgBox ("先查找学生是否存在！")
      ElseIf flag = 0 Then
     MsgBox ("该生不存在")
     Else
   sql = "update member set name='" + Text8.Text + "',sex='" + Text9.Text + "' where id='" + Text7.Text + "'"
   conn.Open connstr
   rs.Open sql, conn, 3, 3
   MsgBox ("已成功更新")
   Adodc1.Refresh
   Set DataGrid1.DataSource = Adodc1
  End If
End Sub
 
Private Sub Command6_Click()                          '删除
Dim conn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   Dim sql As String
   If flag = 2 Then
     MsgBox ("先查询该生是否存在")
    ElseIf flag = 0 Then
     MsgBox ("该生不存在")
     Else
     
   sql = " delete from member where id='" + Text7.Text + "'"
   conn.Open connstr
   rs.Open sql, conn, 3, 2
    MsgBox ("已成功删除")
   Adodc1.Refresh
   Set DataGrid1.DataSource = Adodc1
   End If
End Sub

Private Sub Command7_Click()                      '按姓名查询
Dim sql As String
sql = "select * from member where name='" + Text2.Text + "'"
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = sql
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

