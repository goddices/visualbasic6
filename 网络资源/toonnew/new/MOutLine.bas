Attribute VB_Name = "OutLine"
Type CUSTOMVERTEX
 postion As D3DVECTOR    '3d position for vertex
 normal1 As D3DVECTOR     'surface normal1
 normal2 As D3DVECTOR     'surface normal2
End Type

Public Type silLine
 vertici() As CUSTOMVERTEX
End Type

Public Type Triangle
 P(2) As D3DVECTOR
 n As D3DVECTOR
End Type

Public CALBSP As New vbbsp
Public bspRan() As Long
Public bsplist() As Long
Public bsplistlen As Long

Public mane As String
Public penmane As String
Public cubo As OneMesh
Public i As Long, j As Long, K As Long
Public D As Single

Public objcen As D3DVECTOR
Public objra As Single

Dim EDGC(2, 1) As Long

Function CreaOutLine(m As D3DXMesh) As silLine
 Dim l As Long

 'get buffer
 Dim verti() As D3DVERTEX
 Dim Indici() As Integer
 ReDim verti(m.GetNumVertices - 1)
 ReDim Indici(m.GetNumFaces * 3 - 1)

 D3DXMeshVertexBuffer8GetData m, 0, Len(verti(0)) * (m.GetNumVertices), 0, verti(0)
 D3DXMeshIndexBuffer8GetData m, 0, Len(Indici(0)) * (m.GetNumFaces * 3), 0, Indici(0)

 ReDim CreaOutLine.vertici(m.GetNumFaces * 6)

 Dim n1 As Long
 Dim n2 As Long
 Dim n0 As Long

 Dim Tris() As Triangle
 ReDim Tris(m.GetNumFaces - 1)

 Dim EdgeUse() As Boolean
 ReDim EdgeUse(m.GetNumFaces - 1, 2)

 For i = 0 To m.GetNumFaces - 1
  n0 = Indici(3 * i + 0)
  n1 = Indici(3 * i + 1)
  n2 = Indici(3 * i + 2)
  Tris(i).P(0).x = verti(n0).x
  Tris(i).P(0).y = verti(n0).y
  Tris(i).P(0).Z = verti(n0).Z
  Tris(i).P(1).x = verti(n1).x
  Tris(i).P(1).y = verti(n1).y
  Tris(i).P(1).Z = verti(n1).Z
  Tris(i).P(2).x = verti(n2).x
  Tris(i).P(2).y = verti(n2).y
  Tris(i).P(2).Z = verti(n2).Z
  Tris(i).n = cross(Subtract(Tris(i).P(2), Tris(i).P(1)), Subtract(Tris(i).P(1), Tris(i).P(0)))
  D3DXVec3Normalize Tris(i).n, Tris(i).n
 Next

 EDGC(0, 0) = 0: EDGC(0, 1) = 1
 EDGC(1, 0) = 1: EDGC(1, 1) = 2
 EDGC(2, 0) = 2: EDGC(2, 1) = 0

 Dim Xrev() As D3DVECTOR2 '2叉树优化
 ReDim Xrev(m.GetNumFaces * 3 - 1)

 For i = 0 To m.GetNumFaces - 1
  Xrev(3 * i + 0).x = V2Tot(Tris(i).P(0), Tris(i).P(1))
  Xrev(3 * i + 1).x = V2Tot(Tris(i).P(1), Tris(i).P(2))
  Xrev(3 * i + 2).x = V2Tot(Tris(i).P(2), Tris(i).P(0))
  Xrev(3 * i + 0).y = 3 * i + 0
  Xrev(3 * i + 1).y = 3 * i + 1
  Xrev(3 * i + 2).y = 3 * i + 2
 Next

 CALBSP.push Xrev

 'bsplistlen
 Dim IndbspRan() As Long
 ReDim bspRan(m.GetNumFaces * 3)
 ReDim IndbspRan(m.GetNumFaces * 3)
 ReDim bsplist(0)
 bsplist(0) = 0
 CALBSP.getdata

 For i = 0 To UBound(bspRan)
  IndbspRan(i) = bspRan(i) Mod 3
  bspRan(i) = Int(bspRan(i) / 3)
 Next

 Dim Tri1 As Triangle, Tri2 As Triangle
 Dim up As Long, down As Long

 For K = 1 To UBound(bsplist)
  up = bsplist(K - 1) + 1: down = bsplist(K)

  For i = up To down - 1
   For j = i + 1 To down
    Tri1 = Tris(bspRan(i)): Tri2 = Tris(bspRan(j))

    If CmpTri(Tri1, Tri2, IndbspRan(i), IndbspRan(j)) Then

     If EdgeUse(bspRan(i), IndbspRan(i)) = False Then
      EdgeUse(bspRan(i), IndbspRan(i)) = True
      EdgeUse(bspRan(j), IndbspRan(j)) = True

      If D3DXVec3Dot(Tri1.n, Tri2.n) < 0.9999 Then
       CreaOutLine.vertici(l).postion = Tri1.P(EDGC(IndbspRan(i), 0))
       CreaOutLine.vertici(l).normal1 = Tri1.n
       CreaOutLine.vertici(l).normal2 = Tri2.n
       If D3DXVec3Dot(Tri1.n, Tri2.n) < D Then _
          D3DXVec3Scale CreaOutLine.vertici(l).normal2, CreaOutLine.vertici(l).normal1, -1
       CreaOutLine.vertici(l + 1) = CreaOutLine.vertici(l)
       CreaOutLine.vertici(l + 1).postion = Tri1.P(EDGC(IndbspRan(i), 1))
       l = l + 2
      End If

     End If
    End If
   Next
  Next
 Next

 For i = 0 To m.GetNumFaces - 1
  For j = 0 To 2
   If EdgeUse(i, j) = False Then
    CreaOutLine.vertici(l).postion = Tris(i).P(EDGC(j, 0))
    CreaOutLine.vertici(l).normal1 = Tris(i).n
    CreaOutLine.vertici(l).normal2 = VScale(Tris(i).n, -1)
    CreaOutLine.vertici(l + 1) = CreaOutLine.vertici(l)
    CreaOutLine.vertici(l + 1).postion = Tris(i).P(EDGC(j, 1))
    l = l + 2
   End If
  Next
 Next

 l = l - 1

 '  Dim subs As Long '去除重复边
 '  For i = 0 To l - 2 Step 2
 '  For j = i + 2 To l Step 2
 '
 '  If CmpVer(CreaOutLine.vertici(i + 1).postion, CreaOutLine.vertici(j).postion) And _
     '          CmpVer(CreaOutLine.vertici(i).postion, CreaOutLine.vertici(j + 1).postion) Then
 '  subs = subs + 2
 '  For K = j To l - 2 Step 2
 '  CreaOutLine.vertici(K) = CreaOutLine.vertici(K + 2)
 '  CreaOutLine.vertici(K + 1) = CreaOutLine.vertici(K + 3)
 '  Next
 '  Exit For
 '  End If
 '  Next
 '  Next
 '  l = l - subs
 '  ReDim Preserve CreaOutLine.vertici(l)

 Dim Vmid As D3DVECTOR
 Dim Findd As Boolean

 Do '细分
  Findd = False
  For i = 0 To l Step 2
   If VLength(Subtract(CreaOutLine.vertici(i).postion, CreaOutLine.vertici(i + 1).postion)) > objra / 4 Then
    Vmid = VScale(Add(CreaOutLine.vertici(i).postion, CreaOutLine.vertici(i + 1).postion), 0.5)
    l = l + 2
    Findd = True
    If UBound(CreaOutLine.vertici) < l Then ReDim Preserve CreaOutLine.vertici(l + 100)
    CreaOutLine.vertici(l - 1) = CreaOutLine.vertici(i)
    CreaOutLine.vertici(l) = CreaOutLine.vertici(i + 1)
    CreaOutLine.vertici(i + 1).postion = Vmid
    CreaOutLine.vertici(l - 1).postion = Vmid
   End If
  Next
 Loop Until Not Findd

 ReDim Preserve CreaOutLine.vertici(l)
End Function

Function CmpTri(Tr1 As Triangle, Tr2 As Triangle, rc1 As Long, rc2 As Long) As Boolean

 If CmpVer(Tr1.P(EDGC(rc1, 0)), Tr2.P(EDGC(rc2, 1))) And CmpVer(Tr1.P(EDGC(rc1, 1)), Tr2.P(EDGC(rc2, 0))) Then
  CmpTri = 1
 End If

End Function
