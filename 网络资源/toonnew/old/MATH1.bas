Attribute VB_Name = "MATH1"
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

'为了代码简洁，部分D3D向量改为具有返回值的函数
Public Function vec3(X As Single, Y As Single, Z As Single) As D3DVECTOR
 vec3.X = X
 vec3.Y = Y
 vec3.Z = Z
End Function

Public Function MakeRect(bottom As Long, top As Long, Left As Long, Right As Long) As RECT
 MakeRect.bottom = bottom
 MakeRect.top = top
 MakeRect.Left = Left
 MakeRect.Right = Right
End Function

Public Function vec2(X As Single, Y As Single) As D3DVECTOR2
 vec2.X = X
 vec2.Y = Y
End Function

Public Function vec3Tovec4(ver1 As D3DVECTOR) As D3DVECTOR4
 vec3Tovec4.X = ver1.X
 vec3Tovec4.Y = ver1.Y
 vec3Tovec4.Z = ver1.Z
 vec3Tovec4.w = 1
End Function

Public Function cross(ver1 As D3DVECTOR, ver2 As D3DVECTOR) As D3DVECTOR
 D3DXVec3Cross cross, ver1, ver2
End Function

Public Function Add(ver1 As D3DVECTOR, ver2 As D3DVECTOR) As D3DVECTOR
 D3DXVec3Add Add, ver1, ver2
End Function

Public Function SubNormalize(ver1 As D3DVECTOR, ver2 As D3DVECTOR) As D3DVECTOR
 D3DXVec3Subtract SubNormalize, ver1, ver2
 D3DXVec3Normalize SubNormalize, SubNormalize
End Function

Public Function Subtract(ver1 As D3DVECTOR, ver2 As D3DVECTOR) As D3DVECTOR
 D3DXVec3Subtract Subtract, ver1, ver2
End Function

Public Function Vmid(ver1 As D3DVECTOR, ver2 As D3DVECTOR) As D3DVECTOR
 D3DXVec3Add Vmid, ver1, ver2
 D3DXVec3Scale Vmid, Vmid, 0.5
End Function

Public Function VParam(ver1 As D3DVECTOR, ver2 As D3DVECTOR, Param As Single) As D3DVECTOR
 D3DXVec3Add VParam, VScale(ver2, Param), VScale(ver1, 1 - Param)
End Function

Public Function VScale(ver1 As D3DVECTOR, s As Single) As D3DVECTOR
 D3DXVec3Scale VScale, ver1, s
End Function

Public Function Dot(ver1 As D3DVECTOR, ver2 As D3DVECTOR) As Single
 Dot = D3DXVec3Dot(ver1, ver2)
 If Dot = 0 Then Dot = 0.0001
End Function

Public Function VLength(ver1 As D3DVECTOR) As Single
 VLength = D3DXVec3Length(ver1)
End Function

Public Function VDst(ver1 As D3DVECTOR, ver2 As D3DVECTOR) As Single
 VDst = D3DXVec3Length(Subtract(ver1, ver2))
End Function

Public Function Normalize(ver1 As D3DVECTOR) As D3DVECTOR
 D3DXVec3Normalize Normalize, ver1
End Function

Function CmpVer(P1 As D3DVECTOR, P2 As D3DVECTOR) As Boolean '比较2向量
 If P1.X = P2.X And P1.Y = P2.Y And P1.Z = P2.Z Then CmpVer = True Else CmpVer = False
End Function
