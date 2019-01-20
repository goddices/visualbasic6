VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6570
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GEt"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const HWND_TOP = 0
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40

Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
 
Private Declare Function SetFocusAPI& Lib "user32" Alias "SetFocus" (ByVal hwnd As Long)
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long _
) As Long


Private hwnd1() As Long
Private i As Integer
Private maxlen As Integer
Private indexFrm As Integer



Private Sub Command1_Click()
 
    i = i + 1
    ReDim Preserve hwnd1(i)
    hwnd1(i) = FindWindow("asktao", vbNullString)
 
    SetWindowText hwnd1(i), CStr(i)
    

End Sub

Private Sub Command2_Click()
maxlen = UBound(hwnd1)
MsgBox maxlen
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    indexFrm = indexFrm + 1
    If indexFrm > maxlen Then indexFrm = 1
    Print indexFrm
    SetWindowPos hwnd1(indexFrm), -1, 0, 0, 0, 0, SWP_SHOWWINDOW + SWP_NOSIZE = &H1
    SetFocusAPI hwnd1(indexFrm)
End Sub
