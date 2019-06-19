Attribute VB_Name = "Module1"
Public connstr As String
Public rs As ADODB.Recordset
Public conn As ADODB.Connection

Sub main()
connstr = "Provider=SQLOLEDB.1;persist security info=true;User ID=login1;passward=123;initial catalog=mem;Data source=DESKTOP-K8KU961"
Form1.Show

End Sub



Public Function selectsql(sql As String) As ADODB.Recordset
Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection
connstr = "provider=SQLOLEDB.1;persist security info=true;User ID=login1;passward=123;initial catalog=mem;Data source=DESKTOP-K8KU961"
conn.Open connstr
rs.CursorLocation = adUseClient
rs.Open Trim$(sql), conn, adOpenDynamic, adLockOptimistic
Set selectsql = rs
End Function
