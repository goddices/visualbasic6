VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Q"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3315
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2400
      Top             =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "click"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "lock"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type POINTAPI
        px As Long
        py As Long
End Type

Private Declare Function WindowFromPoint Lib "user32" ( _
    ByVal xPoint As Long, _
    ByVal yPoint As Long _
) As Long

Private Declare Function GetCursorPos Lib "user32" ( _
    lpPoint As POINTAPI _
) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any _
) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Private Declare Function SetCursorPos Lib "user32" ( _
    ByVal x As Long, _
    ByVal y As Long _
) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hwnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long _
) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As Long

Private Const WM_LBUTTONDOWN = &H201
Private Const WM_KEYDOWN = &H100
Private Const VK_RETURN = &HD
Private Const WM_KEYUP = &H101
Private Const WM_LBUTTONUP = &H202
Private Const WM_SETFOCUS = &H7

Dim hwindow As Long
Dim pt As POINTAPI
Dim str As String

Private Sub Command1_Click()
GetCursorPos pt
hwindow = WindowFromPoint(pt.px, pt.py)
 
PostMessage hwindow, WM_LBUTTONDOWN, 0, 0
PostMessage hwindow, WM_LBUTTONUP, 0, 0


str = String(255, 0)

Timer1.Enabled = True

End Sub

Private Sub Command2_Click()
Print "df"
End Sub

Private Sub Timer1_Timer()
Dim h1 As Long, h2 As Long
    SetCursorPos pt.px, pt.py
    
  
    
  
        PostMessage hwindow, WM_LBUTTONDOWN, 0, 0
        PostMessage hwindow, WM_LBUTTONUP, 0, 0
        
        h1 = FindWindow(vbNullString, "QQ对战平台")
        SendMessage h1, WM_SETFOCUS, 0, 0
        PostMessage h1, WM_KEYDOWN, VK_RETURN, 0
        PostMessage h1, WM_KEYUP, VK_RETURN, 0
    
        h2 = FindWindow(vbNullString, "错误")
        SendMessage h1, WM_SETFOCUS, 0, 0
        PostMessage hwindow, WM_KEYDOWN, VK_RETURN, 0
        PostMessage hwindow, WM_KEYUP, VK_RETURN, 0
 
End Sub
