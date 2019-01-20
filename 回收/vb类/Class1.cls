VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 
 Private str As String
 Private sql As String
 Public Sub SetConn(connobj As ADODB.Connection)
 
 str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\web\ado\Ã‚ø‚.mdb"
 
 connobj.Open str
 
 
 End Sub


Public Sub SetRecordSet(rsobj As ADODB.Recordset, connobj As ADODB.Connection)
 

sql = "select * from type2"

rsobj.Open sql, connobj, adOpenStatic, adLockOptimistic
 
End Sub
', conn, adOpenStatic, adLockOptimistic
