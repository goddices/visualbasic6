Attribute VB_Name = "Module1"
Public Type POINTAPI
        dx As Long
        dy As Long
End Type


Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long



Public Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Const WM_LBUTTONDOWN = &H201
Public Const WM_CHAR = &H102

Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101

Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSCHAR = &H106

Public Const VK_RETURN = &HD
Public Const VK_UP = &H26

Public Const VK_MENU = &H12

Public MaxTime As Integer

Public Function Ttime() As Boolean
Static min As Integer
Static sec As Integer
sec = sec + 1
If sec >= 60 Then sec = 0: min = min + 1
If min >= MaxTime Then min = 0: Ttime = True
End Function

