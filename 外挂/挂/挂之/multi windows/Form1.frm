VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6060
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "寻找窗口"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOP = 0
Private Const HWND_BOTTOM = 1

Private Const WM_ACTIVATE = &H6
Private Const WM_SETFOCUS = &H7

Dim hwnd1 As Long
Dim hwnd2 As Long
Dim hWndWenDao() As Long
Dim wndIndex As Integer
Dim i As Integer
Private Sub Command1_Click()
    hwnd1 = FindWindow("AskTao", vbNullString)
    i = i + 1
    ReDim Preserve hWndWenDao(i) As Long
    hWndWenDao(i) = hwnd1
    SetWindowText hwnd1, CStr(i)
    SetWindowPos hwnd1, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE

End Sub

    'SetWindowPos Me.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    'SetWindowPos hwnd1, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    'hwnd2 = FindWindow(vbNullString, "123.txt - 记事本")
    'SetWindowText hwnd2, "321"
    'SetWindowPos hwnd2, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    'SetWindowPos hwnd2, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    
Private Sub Command2_Click()

Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
    Dim aa As Long
    aa = FindWindow("AskTao", vbNullString)
    SetWindowText aa, "dfdfdf'"
End Sub

Private Sub Command4_Click()
    For n = 1 To UBound(hWndWenDao)
        SetWindowPos hWndWenDao(hWndWenDao), HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE
    Next
End Sub

Private Sub Timer1_Timer()
    wndIndex = wndIndex + 1
    If wndIndex > UBound(hWndWenDao) Then wndIndex = 1
    'SetWindowPos hWndWenDao(wndIndex - 1), HWND_BOTTOM, (wndIndex - 1) * 100, (wndIndex - 1) * 100, 0, 0, SWP_NOSIZE
   
    SetWindowPos hWndWenDao(wndIndex), HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE
     SendMessage hWndWenDao(wndIndex), WM_SETFOCUS, 0, 0
    'SetWindowPos hWndWenDao(wndIndex), HWND_BOTTOM, (wndIndex) * 100, (wndIndex) * 100, 0, 0, SWP_NOSIZE
    Sleep 100
 
End Sub
