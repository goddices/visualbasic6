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



Public Const deltaWidth As Integer = 20
Public Const deltaHeight  As Integer = deltaWidth
Public Const intMaxX As Integer = 15
Public Const intMaxY As Integer = 15


Public Type RECT
    startX As Long
    startY As Long
    endX As Long
    endY As Long
End Type

Public Type POINT
    ptX As Integer
    ptY As Integer
End Type

Public Type BLOCK
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
End Type

Public Range As RECT
Public hBrush As Long
Public intCoordinates(-1 To intMaxX, -1 To intMaxY) As Integer
Public mvarMainPictureBox As PictureBox

Public o(2) As POINT 'Object 1

'0 Empty
'1 Walls And Bounds
'2 Controller
'3 Objects
'4 Targets

