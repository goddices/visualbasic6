Attribute VB_Name = "Module1"
Public Conn As New ADODB.Connection
Public Rs As New ADODB.Recordset


Public Sub OpenDateBase()
Set Conn = New ADODB.Connection
Set Rs = New ADODB.Recordset
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\login.mdb"
Rs.Open "select * from pwd", Conn, adOpenStatic, adLockOptimistic

End Sub


Public Sub Username()

End Sub
