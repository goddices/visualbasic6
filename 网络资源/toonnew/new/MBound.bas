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
Public Varray2() As TypeLine
Public arraylen  As Long
Public schain() As New vbchain

 Public nestdata() As D3DVECTOR2
 
Dim i As Long, j As Long

Public Sub AddToArray(ver1 As D3DVECTOR, ver2 As D3DVECTOR)
 arraylen = arraylen + 1 '0不用
 Varray(arraylen).P1 = ver1
 Varray(arraylen).P2 = ver2
End Sub



Public Function FindAndDrawChian()
 Dim i As Long, j As Long, CurC As Long
 ReDim schain(0)
If arraylen = 0 Then Exit Function

Set CALBSP = Nothing
Set CALBSP = New vbbsp
 Dim Xrev() As D3DVECTOR2 '2叉树优化
 ReDim Xrev(Abs(arraylen * 2 - 1))
 
 For i = 0 To arraylen - 1
  Xrev(2 * i + 0).x = Varray(i + 1).P1.x
  Xrev(2 * i + 1).x = Varray(i + 1).P2.x
  Xrev(2 * i + 0).y = i + 1
  Xrev(2 * i + 1).y = i + 1
 Next

 CALBSP.push Xrev


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
   
 ReDim nestdata(0)
 CALBSP.getEqudata (schain(CurC).fon.x)
   For i = 0 To UBound(nestdata)
    If Varray(nestdata(i).y).use = 0 Then
     If schain(CurC).push(Varray(nestdata(i).y).P1, Varray(nestdata(i).y).P2) Then
   Varray(nestdata(i).y).use = 1: findchainnode = 1: Exit For
    End If
    End If
   Next

 ReDim nestdata(0)
 CALBSP.getEqudata (schain(CurC).bac.x)
   For i = 0 To UBound(nestdata)
    If Varray(nestdata(i).y).use = 0 Then
     If schain(CurC).push(Varray(nestdata(i).y).P1, Varray(nestdata(i).y).P2) Then
    Varray(nestdata(i).y).use = 1: findchainnode = 1: Exit For
    End If
    End If
   Next
   
  Loop Until Not findchainnode
 Loop Until Not findchain


 For i = 0 To UBound(schain)
  schain(i).DEC
  drawChain schain(i).getdata
 Next
End Function
