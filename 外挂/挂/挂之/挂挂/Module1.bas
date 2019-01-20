Attribute VB_Name = "Module1"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONDBLCLK = &H203
Public Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Public Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Public Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Public Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
Public Const HWND_TOPMOST = -1
Public Const SWP_SHOWWINDOW = &H40
Public Const SW_RESTORE = 9
Public Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move

Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1

Public Const WM_ACTIVATE = &H6
Public Const WM_SETFOCUS = &H7


Public nIndex As Long
 
Public hWD() As Long
Public mCount As Integer
Public rect2 As RECT
Public pt As POINTAPI

Public Function ProcFunc(ByVal hwnd As Long, ByVal lParam As Long) As Long
  
    Dim strClsName As String
    Dim strWndTxt As String
    strClsName = Space(255)
    strWndTxt = Space(255)
    If hwnd <> 0 Then
         
        'ReDim Preserve strClsName(nIndex) As String
        GetClassName hwnd, strClsName, 255
        GetWindowText hwnd, strWndTxt, 255
        
        
        'ThunderRT6FormDC
        'AskTao
        If Left(strClsName, Len("AskTao")) = "AskTao" Then
            nIndex = nIndex + 1
            ReDim Preserve hWD(nIndex) As Long
            hWD(nIndex) = hwnd
            SetWindowText hwnd, "SB" & CStr(nIndex) & "ºÅ"
            ProcFunc = 1
        End If
        ProcFunc = 1
    Else
        ProcFunc = 0
    End If
End Function

Public Function GetHandle() As Integer
    If (EnumWindows(AddressOf ProcFunc, 0)) Then
        GetHandle = 1
    Else
        GetHandle = 0
    End If
End Function

Public Sub click()
    mCount = mCount + 1
    If mCount > UBound(hWD) Then mCount = 1
       
    GetCursorPos pt
    ShowWindow hWD(mCount), SW_RESTORE
    SetWindowPos hWD(mCount), HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE
    GetWindowRect hWD(mCount), rect2
    PostMessage hWD(mCount), WM_SETFOCUS, 0, 0
    
    SetCursorPos rect2.Left + 100, rect2.Top + 10
    mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Sleep 200
    
    SetCursorPos rect2.Left + 600, rect2.Top + 60
    'PostMessage hwnd2, WM_RBUTTONDOWN, 0, 0
    'PostMessage hwnd2, WM_RBUTTONUP, 0, 0
    mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    Sleep 1000
  
    
    SetCursorPos rect2.Left + 600, rect2.Top + 78
    'PostMessage hwnd2, WM_RBUTTONDOWN, 0, 0
    'PostMessage hwnd2, WM_RBUTTONUP, 0, 0
    mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    Sleep 1000
    
    SetCursorPos rect2.Left + 740, rect2.Top + 60
    'PostMessage hwnd2, WM_RBUTTONDOWN, 0, 0
    'PostMessage hwnd2, WM_RBUTTONUP, 0, 0
    mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    Sleep 1000
    
    SetCursorPos rect2.Left + 740, rect2.Top + 78
    'PostMessage hwnd2, WM_RBUTTONDOWN, 0, 0
    'PostMessage hwnd2, WM_RBUTTONUP, 0, 0
    mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    Sleep 1000
    
  
     
    SetCursorPos pt.x, pt.y

End Sub
