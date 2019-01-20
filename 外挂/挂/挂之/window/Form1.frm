VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   915
   ClientLeft      =   12000
   ClientTop       =   1020
   ClientWidth     =   2790
   LinkTopic       =   "Form1"
   ScaleHeight     =   915
   ScaleWidth      =   2790
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   4320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hwnd1 As Long
Dim rect1 As RECT
Dim pt As POINTAPI

'Dim str As String
'str = "Left:     " & rect1.Left & vbNewLine & _
      "Top:       " & rect1.Top & vbNewLine & _
      "Right:   " & rect1.Right & vbNewLine & _
      "Bottom: " & rect1.Bottom & vbNewLine & _
      "Width:   " & CStr(rect1.Right - rect1.Left) & vbNewLine & _
      "Height: " & CStr(rect1.Bottom - rect1.Top)
'MsgBox str
'MsgBox hwnd1
'Dim gwnd As Long
'gwnd = GetWindowLong(hwnd1, GWL_STYLE)
'MsgBox gwnd


Private Sub Command1_Click()
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    Dim hwnd2 As Long
    Dim rect2 As RECT
    hwnd2 = FindWindow("asktao", vbNullString)
                                     'Î´ÃüÃû.bmp - Picasa ÕÕÆ¬²é¿´Æ÷
    GetWindowRect hwnd2, rect2
    'MsgBox hwnd2
    Dim str As String
    'str = "Left:     " & rect2.Left & vbNewLine & _
         "Top:       " & rect2.Top & vbNewLine & _
         "Right:   " & rect2.Right & vbNewLine & _
         "Bottom: " & rect2.Bottom & vbNewLine & _
         "Width:   " & CStr(rect2.Right - rect2.Left) & vbNewLine & _
         "Height: " & CStr(rect2.Bottom - rect2.Top)
    'MsgBox str
 '740    60    78
 '600    60
    SetWindowPos hwnd2, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE
     
    SetCursorPos rect2.Left + 600, rect2.Top + 60
    PostMessage hwnd2, WM_RBUTTONDOWN, 0, 0
    PostMessage hwnd2, WM_RBUTTONUP, 0, 0
   ' mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
   ' mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    Sleep 100
    
    SetCursorPos rect2.Left + 600, rect2.Top + 78
    PostMessage hwnd2, WM_RBUTTONDOWN, 0, 0
    PostMessage hwnd2, WM_RBUTTONUP, 0, 0
   ' mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
   ' mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
     Sleep 100
    
    SetCursorPos rect2.Left + 740, rect2.Top + 60
    PostMessage hwnd2, WM_RBUTTONDOWN, 0, 0
    PostMessage hwnd2, WM_RBUTTONUP, 0, 0
    ' mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
   ' mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    Sleep 100
    
    SetCursorPos rect2.Left + 740, rect2.Top + 78
    PostMessage hwnd2, WM_RBUTTONDOWN, 0, 0
    PostMessage hwnd2, WM_RBUTTONUP, 0, 0
   ' mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
   ' mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0

    Sleep 100
    
    PostMessage hwnd2, WM_RBUTTONDOWN, 0, 0
    PostMessage hwnd2, WM_RBUTTONUP, 0, 0
End Sub

Private Sub Form_Load()

    hwnd1 = FindWindow("asktao", vbNullString)
    GetWindowRect hwnd1, rect1
    If hwnd1 = 0 Then MsgBox "window not open!"
    'SetWindowPos hwnd1, HWND_TOPMOST, 0, 0, 1024, 768, SWP_SHOWWINDOW
End Sub

Private Sub Timer1_Timer()
    GetWindowRect hwnd1, rect1
    GetCursorPos pt
    
    SetCursorPos rect1.Left + 500, rect1.Top + 60
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
    SetCursorPos rect1.Left + 650, rect1.Top + 60
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    SetCursorPos pt.x, pt.y
End Sub
