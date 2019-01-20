VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WM_ACTIVATE = &H6
Private Const WM_SETFOCUS = &H7
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_RESTORE = 9

Dim mhwnd(2) As Long

Private Sub Command1_Click()

mhwnd(0) = FindWindow(vbNullString, "我的电脑")
mhwnd(1) = FindWindow(vbNullString, "我的文档")
mhwnd(2) = FindWindow(vbNullString, "回收站")
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Static i As Integer
i = i + 1
If i > 2 Then i = 0
ShowWindow mhwnd(i), SW_RESTORE
SetWindowPos mhwnd(i), -1, 0, 0, 0, 0, 1
SendMessage mhwnd(i), WM_SETFOCUS, 0, 0
End Sub
