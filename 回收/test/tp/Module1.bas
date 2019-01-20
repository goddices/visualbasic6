Attribute VB_Name = "Module1"
Public Declare Function _
Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" ( _
    ByVal dwMessage As Long, _
    lpData As NOTIFYICONDATA _
) As Long
'//bullshit
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
) As Long


Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long _
) As Long

Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
Public Const WM_CLOSE = &H10
Public Const WM_USER = &H400
Public Const GWL_WNDPROC = (-4)


Public Const TRAY_CALLBACK = (WM_USER + 1001&)

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

Public nid As NOTIFYICONDATA

Public OldWindowProc   As Long

Public Sub Shell_NothifyIcon_Example()
    
    OldWindowProc = SetWindowLong(Form1.hwnd, GWL_WNDPROC, AddressOf WndProc)
    
    nid.cbSize = Len(nid)
    nid.hIcon = Form1.Icon.Handle
    nid.hwnd = Form1.hwnd
    nid.szTip = vbNullString
    nid.uCallbackMessage = TRAY_CALLBACK
    nid.uFlags = NIF_ICON
    nid.uID = 100
   
    
    Call Shell_NotifyIcon(NIM_ADD, nid)
End Sub

Private Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
 
    Form1.Print Msg
    If Msg = TRAY_CALLBACK Then
       
        If lParam = WM_LBUTTONUP Then
            Form1.PopupMenu dd
            If Form1.WindowState = vbMinimized Then _
                  
                Form1.WindowState = 3
            Form1.SetFocus
             
            End If
        End If
        '如果点击了右键
        
    Else
    
        '如果是其他类型的消息则传递给原有默认的窗口函数
        WndProc = CallWindowProc(OldWindowProc, hwnd, Msg, wParam, lParam)
    End If
End Function

Public Function WndProc222(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
        Case WM_CLOSE
            Unload Me
        
    End Select
End Function

Public Function dd() As Long
    dd = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf WndProc222)
    callwindowproc(
End Function
