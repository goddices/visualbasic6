Attribute VB_Name = "Module1"


Public Declare Function CreateSolidBrush Lib "gdi32" ( _
    ByVal crColor As Long _
) As Long


Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function FillRect Lib "user32" ( _
    ByVal hdc As Long, _
    lpRect As RECT, _
    ByVal hBrush As Long _
) As Long



Public Const deltaWidth = 20
Public Const deltaHeight = deltaWidth
Public Const intMaxX As Integer = 15
Public Const intMaxY As Integer = 10


Public Type RECT
    startX As Long
    startY As Long
    endX As Long
    endY As Long
End Type


Public Range As RECT
Public hBrush As Long
Public intCoordinates(-1 To intMaxX, -1 To intMaxY) As Integer

Public mvarMainPictureBox As Object        '局部复制
Public mvarSecondaryPictureBox As Object '局部复制


Public Sub Error_001()
        MsgBox "error 001: main picture box is nothing", vbOKOnly + vbExclamation
End Sub

Public Sub Error_002()
    MsgBox "error 002: target object is not a picture box", vbOKOnly + vbExclamation
End Sub
