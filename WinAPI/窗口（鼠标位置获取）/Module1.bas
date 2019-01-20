Attribute VB_Name = "Module1"
Public Type POINTAPI
        dx As Long
        dy As Long
End Type


Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

