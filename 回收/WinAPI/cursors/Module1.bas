Attribute VB_Name = "Module1"
Public Type POINTAPI
    sx As Long
    sy As Long
End Type

Public Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type


Public Declare Sub mouse_event Lib "user32" ( _
    ByVal dwFlags As Long, _
    ByVal dx As Long, _
    ByVal dy As Long, _
    ByVal cButtons As Long, _
    ByVal dwExtraInfo As Long _
)

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" ( _
    ByVal hwnd As Long, _
    ByVal lpText As String, _
    ByVal lpCaption As String, _
    ByVal wType As Long _
) As Long




Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" ( _
    lpMsg As MSG, _
    ByVal hwnd As Long, _
    ByVal wMsgFilterMin As Long, _
    ByVal wMsgFilterMax As Long _
) As Long

Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)

Public Const MB_OK = &H0&
Public Const MB_ICONINFORMATION = &H40&
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSELAST = &H209
Public Const WM_DESTROY = &H2



Public Point As POINTAPI
Public uMsg As MSG

Sub Main() 'App.hInstance
Form1.Show
Call GetCursorPos(Point)
MessageBox Form1.hwnd, CStr(Point.sx & " " & Point.sy), "ddd", MB_OK + MB_ICONINFORMATION

End Sub


