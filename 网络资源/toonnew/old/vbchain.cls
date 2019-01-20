VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbchain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Á´±í(VB)
Dim data() As D3DVECTOR
Dim fornt As D3DVECTOR
Dim back As D3DVECTOR
Dim LS As Long
Dim j As Long

Public Function Testa(V1 As D3DVECTOR, V2 As D3DVECTOR, V3 As D3DVECTOR) As Boolean
Const ANGLIM As Single = -0.1 '·§Öµ
If Dot(SubNormalize(V1, V2), SubNormalize(V3, V2)) < ANGLIM Then Testa = True
End Function

Public Function push(P1 As D3DVECTOR, P2 As D3DVECTOR) As Boolean
If CmpVer(fornt, P1) And Testa(P2, data(LS), data(LS - 1)) Then
pushfornt P2: push = 1
ElseIf CmpVer(fornt, P2) And Testa(P1, data(LS), data(LS - 1)) Then
pushfornt P1: push = 1
ElseIf CmpVer(back, P1) And Testa(P2, data(1), data(2)) Then
pushback P2: push = 1
ElseIf CmpVer(back, P2) And Testa(P1, data(1), data(2)) Then
pushback P1: push = 1
End If
End Function

Public Sub pushback(V1 As D3DVECTOR)
LS = LS + 1
ReDim Preserve data(LS)
For j = LS To 1 Step -1
data(j) = data(j - 1)
Next
data(1) = V1
back = V1
End Sub

Public Sub pushfornt(V1 As D3DVECTOR)
LS = LS + 1
ReDim Preserve data(LS)
data(LS) = V1
fornt = V1
End Sub

Public Sub DEC()
LS = LS - 1
For j = 0 To LS
data(j) = data(j + 1)
Next
If LS >= 0 Then ReDim Preserve data(LS)
End Sub

Private Sub Class_Initialize()
ReDim data(0)
End Sub

Public Sub clears()
ReDim data(0)
LS = 0
End Sub
Public Property Get fon() As D3DVECTOR
fon = fornt
End Property

Public Property Let fon(vNewValue As D3DVECTOR)
fornt = vNewValue
End Property

Public Property Get bac() As D3DVECTOR
bac = back
End Property

Public Property Let bac(vNewValue As D3DVECTOR)
back = vNewValue
End Property

Public Function getdata() As D3DVECTOR()
getdata = data
End Function
