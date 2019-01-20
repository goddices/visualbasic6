Attribute VB_Name = "OutLine"
Type CUSTOMVERTEX
postion As D3DVECTOR    '3d position for vertex
normal1 As D3DVECTOR     'surface normal1
normal2 As D3DVECTOR     'surface normal2
End Type

Public mane As String
Public cubo As OneMesh
Public i As Long, j As Long, K As Long
Public D As Single

Public objcen As D3DVECTOR
Public objra As Single

Public Type ombra
vertici() As CUSTOMVERTEX
End Type


Function CreaOutLine(m As D3DXMesh) As ombra

Dim l As Long
Dim vertici() As D3DVERTEX
Dim Indici() As Integer
ReDim vertici(m.GetNumVertices - 1)
ReDim Indici(m.GetNumFaces * 3 - 1)

Dim verc() As D3DVECTOR
Dim norma() As D3DVECTOR
Dim norma2() As D3DVECTOR
ReDim verc(m.GetNumFaces * 3 - 1)
ReDim norma(m.GetNumFaces - 1)
ReDim norma2(m.GetNumFaces - 1)

Dim pos() As D3DVECTOR
ReDim pos(m.GetNumFaces * 3)

ReDim CreaOutLine.vertici(m.GetNumFaces * 6 - 1)

'i buffer
D3DXMeshVertexBuffer8GetData m, 0, Len(vertici(0)) * (m.GetNumVertices), 0, vertici(0)
D3DXMeshIndexBuffer8GetData m, 0, Len(Indici(0)) * (m.GetNumFaces * 3), 0, Indici(0)

Dim n1 As Long
Dim n2 As Long
Dim n0 As Long

Dim v0 As D3DVECTOR
Dim V1 As D3DVECTOR
Dim V2 As D3DVECTOR

Dim vettore As D3DVECTOR
Dim vettore1 As D3DVECTOR
Dim vettore2 As D3DVECTOR

For i = 0 To m.GetNumFaces - 1
n0 = Indici(3 * i + 0)
n1 = Indici(3 * i + 1)
n2 = Indici(3 * i + 2)
verc(3 * i + 0).X = vertici(n0).X
verc(3 * i + 0).Y = vertici(n0).Y
verc(3 * i + 0).Z = vertici(n0).Z
verc(3 * i + 1).X = vertici(n1).X
verc(3 * i + 1).Y = vertici(n1).Y
verc(3 * i + 1).Z = vertici(n1).Z
verc(3 * i + 2).X = vertici(n2).X
verc(3 * i + 2).Y = vertici(n2).Y
verc(3 * i + 2).Z = vertici(n2).Z
Next

For i = 0 To m.GetNumFaces - 1
v0 = verc(3 * i + 0)
V1 = verc(3 * i + 1)
V2 = verc(3 * i + 2)
D3DXVec3Subtract vettore1, V2, V1
D3DXVec3Subtract vettore2, V1, v0
D3DXVec3Cross vettore, vettore1, vettore2
D3DXVec3Normalize vettore, vettore
norma(i) = vettore
D3DXVec3Scale norma2(i), norma(i), -1
Next



For i = 0 To m.GetNumFaces - 1

v0 = verc(3 * i + 0)
V1 = verc(3 * i + 1)
V2 = verc(3 * i + 2)

CreaOutLine.vertici(l).normal1 = norma(i)
CreaOutLine.vertici(l + 1).normal1 = norma(i)
CreaOutLine.vertici(l).postion = v0
CreaOutLine.vertici(l + 1).postion = V1
CreaOutLine.vertici(l).normal2 = norma2(i)
CreaOutLine.vertici(l + 1).normal2 = norma2(i)
For j = 0 To m.GetNumFaces - 1
If (CmpVer(verc(3 * i + 0), verc(3 * j + 1)) And CmpVer(verc(3 * i + 1), verc(3 * j + 0))) _
Or (CmpVer(verc(3 * i + 0), verc(3 * j + 2)) And CmpVer(verc(3 * i + 1), verc(3 * j + 1))) _
Or (CmpVer(verc(3 * i + 0), verc(3 * j + 0)) And CmpVer(verc(3 * i + 1), verc(3 * j + 2))) _
Then
If D3DXVec3Dot(norma(j), norma(i)) > D Then
CreaOutLine.vertici(l).normal2 = norma(j)
CreaOutLine.vertici(l + 1).normal2 = norma(j)
Exit For
End If
End If
Next


l = l + 2
CreaOutLine.vertici(l).normal1 = norma(i)
CreaOutLine.vertici(l + 1).normal1 = norma(i)
CreaOutLine.vertici(l).postion = V1
CreaOutLine.vertici(l + 1).postion = V2
CreaOutLine.vertici(l).normal2 = norma2(i)
CreaOutLine.vertici(l + 1).normal2 = norma2(i)
For j = 0 To m.GetNumFaces - 1
If (CmpVer(verc(3 * i + 1), verc(3 * j + 1)) And CmpVer(verc(3 * i + 2), verc(3 * j + 0))) _
Or (CmpVer(verc(3 * i + 1), verc(3 * j + 2)) And CmpVer(verc(3 * i + 2), verc(3 * j + 1))) _
Or (CmpVer(verc(3 * i + 1), verc(3 * j + 0)) And CmpVer(verc(3 * i + 2), verc(3 * j + 2))) _
Then
If D3DXVec3Dot(norma(j), norma(i)) > D Then
CreaOutLine.vertici(l).normal2 = norma(j)
CreaOutLine.vertici(l + 1).normal2 = norma(j)
Exit For
End If
End If
Next



l = l + 2
CreaOutLine.vertici(l).normal1 = norma(i)
CreaOutLine.vertici(l + 1).normal1 = norma(i)
CreaOutLine.vertici(l).postion = V2
CreaOutLine.vertici(l + 1).postion = v0
CreaOutLine.vertici(l).normal2 = norma2(i)
CreaOutLine.vertici(l + 1).normal2 = norma2(i)
For j = 0 To m.GetNumFaces - 1
If (CmpVer(verc(3 * i + 2), verc(3 * j + 1)) And CmpVer(verc(3 * i + 0), verc(3 * j + 0))) _
Or (CmpVer(verc(3 * i + 2), verc(3 * j + 2)) And CmpVer(verc(3 * i + 0), verc(3 * j + 1))) _
Or (CmpVer(verc(3 * i + 2), verc(3 * j + 0)) And CmpVer(verc(3 * i + 0), verc(3 * j + 2))) _
Then
If D3DXVec3Dot(norma(j), norma(i)) > D Then
CreaOutLine.vertici(l).normal2 = norma(j)
CreaOutLine.vertici(l + 1).normal2 = norma(j)
Exit For
End If
End If
Next
l = l + 2
DoEvents
Next
l = l - 1

Dim subs As Long 'È¥³ýÖØ¸´±ß
For i = 0 To l - 2 Step 2
For j = i + 2 To l Step 2

If CmpVer(CreaOutLine.vertici(i + 1).postion, CreaOutLine.vertici(j).postion) And _
CmpVer(CreaOutLine.vertici(i).postion, CreaOutLine.vertici(j + 1).postion) Then
subs = subs + 2
For K = j To l - 2 Step 2
CreaOutLine.vertici(K) = CreaOutLine.vertici(K + 2)
CreaOutLine.vertici(K + 1) = CreaOutLine.vertici(K + 3)
Next
Exit For
End If
Next
Next
l = l - subs
ReDim Preserve CreaOutLine.vertici(l)

Dim Vmid As D3DVECTOR
Dim Findd As Boolean

Do
Findd = False
For i = 0 To l Step 2 'Ï¸·Ö
If VLength(Subtract(CreaOutLine.vertici(i).postion, CreaOutLine.vertici(i + 1).postion)) > objra / 4 Then
Vmid = VScale(Add(CreaOutLine.vertici(i).postion, CreaOutLine.vertici(i + 1).postion), 0.5)
l = l + 2
Findd = True
If UBound(CreaOutLine.vertici) < l Then ReDim Preserve CreaOutLine.vertici(l + 20)
CreaOutLine.vertici(l - 1) = CreaOutLine.vertici(i)
CreaOutLine.vertici(l) = CreaOutLine.vertici(i + 1)
CreaOutLine.vertici(i + 1).postion = Vmid
CreaOutLine.vertici(l - 1).postion = Vmid
End If
Next
Loop Until Not Findd

ReDim Preserve CreaOutLine.vertici(l)
End Function

