VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbbsp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'2²æÊ÷(VB)
Dim data() As D3DVECTOR2

Dim MaxVal As Single
Dim MinVal As Single
Dim MidVal As Single

Dim LS As Long
Dim islef As Boolean
Dim deep As Long

Dim Lchild As vbbsp             ' child vbbsp
Dim Rchild As vbbsp             ' Child vbbsp

Public Function push(XVal() As D3DVECTOR2, Optional dep As Long = 0) As Boolean
 Dim i As Long, j As Long

 deep = dep
 MinVal = XVal(0).x
 MaxVal = XVal(0).x
 LS = UBound(XVal)

 For i = 0 To LS
  If XVal(i).x > MaxVal Then MaxVal = XVal(i).x
  If XVal(i).x < MinVal Then MinVal = XVal(i).x
 Next

 If (MaxVal - MinVal) < 0.00002 Or deep > 10 Or LS = 0 Then
  islef = True
  data = XVal
  Exit Function
 End If

 MidVal = (MaxVal + MinVal) / 2

 Dim Ldat() As D3DVECTOR2
 Dim Rdat() As D3DVECTOR2
 Dim Lls As Long
 Dim Rls As Long
 ReDim Ldat(LS)
 ReDim Rdat(LS)

 For i = 0 To LS
  If XVal(i).x > MidVal Then
   Rdat(Rls) = XVal(i)
   Rls = Rls + 1
  Else
   Ldat(Lls) = XVal(i)
   Lls = Lls + 1
  End If
 Next

 If Rls <> O Then
  ReDim Preserve Rdat(Rls - 1)
  Set Rchild = New vbbsp
  Call Rchild.push(Rdat, deep + 1)
 End If

 If Lls <> O Then
  ReDim Preserve Ldat(Lls - 1)
  Set Lchild = New vbbsp
  Call Lchild.push(Ldat, deep + 1)
 End If

End Function

Public Function getEqudata(XVal As Single) As D3DVECTOR2()
 If islef Then
  If XVal <= MaxVal And XVal >= MinVal Then nestdata = data
 Else
  If XVal > MidVal Then
   If (Not Rchild Is Nothing) Then Rchild.getEqudata (XVal)
  Else
   If (Not Lchild Is Nothing) Then Lchild.getEqudata (XVal)
  End If
 End If
End Function

Public Sub getdata()
 Dim i As Long
 Dim l As Long
 If islef Then
  l = UBound(data)
  bsplistlen = bsplistlen + l + 1
  ReDim Preserve bsplist(UBound(bsplist) + 1)
  bsplist(UBound(bsplist)) = bsplistlen
  For i = 0 To l
   bspRan(bsplistlen - l + i) = data(i).y
  Next
 End If

 If (Not Lchild Is Nothing) Then Lchild.getdata
 If (Not Rchild Is Nothing) Then Rchild.getdata
End Sub

Private Sub Class_Initialize()
 Dim data(0)
End Sub

Private Sub Class_Terminate()
 If deep = 0 Then
  bsplistlen = 0
  ReDim bsplist(O)
  ReDim bspRan(0)
 End If
 Set Lchild = Nothing
 Set Rchild = Nothing
End Sub
