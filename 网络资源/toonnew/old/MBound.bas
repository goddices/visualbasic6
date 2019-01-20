Attribute VB_Name = "MBounds"
'边界查找
Public T As Direct3DTexture8

Public Type CUSTOMVERCOL
 P As D3DVECTOR
 color As Long       'vertex color
End Type

Public Type TypeLine
 P1 As D3DVECTOR
 P2 As D3DVECTOR
 use As Long
End Type

Public Const D3DFVF_CUSTOMVERCOL = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)
Dim findchainnode As Boolean
Dim findchain As Boolean

Public Varray() As TypeLine
Dim arraylen  As Long
Dim schain() As New vbchain

Dim i As Long, j As Long

Public Sub AddToArray(Vis() As CUSTOMVERCOL)
 arraylen = UBound(Varray) + 1 '0不用
 ReDim Preserve Varray(arraylen)
 Varray(arraylen).P1 = Vis(0).P
 Varray(arraylen).P2 = Vis(1).P
End Sub

Public Sub DelArray(Vis() As CUSTOMVERCOL)
 For i = 1 To UBound(Varray) '0不用
  If CmpVer(Varray(i).P1, Vis(0).P) And CmpVer(Varray(i).P2, Vis(1).P) Then
   For j = i To UBound(Varray) - 1
    Varray(j) = Varray(j + 1)
   Next
   ReDim Preserve Varray(UBound(Varray) - 1)
   Exit For
  End If
 Next

End Sub

Public Function FindAndDrawChian()
 Dim i As Long, j As Long, CurC As Long
 ReDim schain(0)
 arraylen = UBound(Varray)

 For i = 1 To arraylen
  Varray(i).P1 = GetScreen(Varray(i).P1)
  Varray(i).P2 = GetScreen(Varray(i).P2)
 Next

 CurC = -1
 Do
  findchain = 0
  For j = 1 To arraylen
   If Varray(j).use = 0 Then
    findchain = 1
    CurC = CurC + 1
    ReDim Preserve schain(CurC) As New vbchain
    Varray(j).use = 1
    schain(CurC).pushfornt Varray(j).P1
    schain(CurC).pushback Varray(j).P2
    Exit For
   End If
  Next

  Do
   findchainnode = False
   For i = 1 To arraylen
    If Varray(i).use = 0 Then
     If schain(CurC).push(Varray(i).P1, Varray(i).P2) Then findchainnode = 1
    End If
    If findchainnode Then Varray(i).use = 1: Exit For
   Next
  Loop Until Not findchainnode

 Loop Until Not findchain


 For i = 0 To UBound(schain)
  schain(i).DEC
  drawChain schain(i).getdata
 Next
End Function
