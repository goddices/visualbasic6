Attribute VB_Name = "Mod_MATH"
Option Explicit

Public Const pi As Single = 3.14159265358979
Public Const PITWO As Single = pi * 2
Public Const PI_DIV2 As Single = pi / 2

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long

'为了代码简洁，部分D3D向量改为具有返回值的函数
Public Function Makever(X As Single, Y As Single, Optional z As Single = 0) As D3DVECTOR
    Makever.X = X
    Makever.Y = Y
    Makever.z = z
End Function

Public Function cross(ver1 As D3DVECTOR, ver2 As D3DVECTOR) As D3DVECTOR
    D3DXVec3Cross cross, ver1, ver2
End Function

Public Function Add(ver1 As D3DVECTOR, ver2 As D3DVECTOR) As D3DVECTOR
    D3DXVec3Add Add, ver1, ver2
End Function

Public Function Subtract(ver1 As D3DVECTOR, ver2 As D3DVECTOR) As D3DVECTOR
    D3DXVec3Subtract Subtract, ver1, ver2
End Function

Public Function VDst(ver1 As D3DVECTOR, ver2 As D3DVECTOR) As Single
    VDst = D3DXVec3Length(Subtract(ver1, ver2))
End Function

Public Function VScale(ver1 As D3DVECTOR, S As Single) As D3DVECTOR
    D3DXVec3Scale VScale, ver1, S
End Function

Public Function Dot(ver1 As D3DVECTOR, ver2 As D3DVECTOR) As Single
    Dot = D3DXVec3Dot(ver1, ver2)
End Function

Public Function VLength(ver1 As D3DVECTOR) As Single
    VLength = D3DXVec3Length(ver1)
End Function

Public Function Normalize(ver1 As D3DVECTOR) As D3DVECTOR
    D3DXVec3Normalize Normalize, ver1
End Function

'Z轴旋转向量
Public Function RotateZ(ver1 As D3DVECTOR, angle As Single) As D3DVECTOR
    Dim X As Single, Y As Single
    X = ver1.X
    Y = ver1.Y
    RotateZ.X = (X) * Cos(angle) - (Y) * Sin(angle)
    RotateZ.Y = (X) * Sin(angle) + (Y) * Cos(angle)
End Function

Public Function RanRnd(Low As Single, Hei As Single) As Single
    RanRnd = Low + (Hei - Low) * Rnd()
End Function

Public Function MaxVel(D1 As Single, D2 As Single)
    If D1 > D2 Then MaxVel = D1 Else MaxVel = D2
End Function


Public Function MinVel(D1 As Single, D2 As Single)
    If D1 < D2 Then MinVel = D1 Else MinVel = D2
End Function

