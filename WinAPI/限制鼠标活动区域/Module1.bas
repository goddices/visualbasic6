Attribute VB_Name = "Module1"
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Public Declare Function ClipCursor Lib "user32" (lpRect As RECT) As Long

Sub main()
Dim rect1 As RECT
rect1.Left = 100
rect1.Right = 1000
rect1.Top = 100
rect1.Bottom = 1000
Call ClipCursor(rect1)
End Sub
