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
Attribute VB_Exposed = True
Public Function abc(ByVal stOperator As Long, ByVal ndOperator As Long, ByVal symbol As Byte)

If (VarType(stOperator) <> vbInteger Or _
    VarType(ndOperator) <> vbInteger Or _
    VarType(symbol) <> vbByte) Then
    
    abc = 0
Else
    abc = stOperator + ndOperator
End If
End Function
