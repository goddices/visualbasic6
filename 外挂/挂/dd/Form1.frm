VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "dd"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   1935
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   -120
      Top             =   1080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "结束"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   -120
      Top             =   600
   End
   Begin VB.Timer TimerFindWindow 
      Interval        =   5000
      Left            =   -120
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "暂停"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'问道(1.411.0508) May 21 15:27:19 2009  [苏堤春晓三线] [asdddddd]
'Counter-Strike
'SendMessage hwnd, WM_KEYDOWN, VK_RETURN, 0


Dim dhwnd As Long

Private Sub Command1_Click()
Timer1.Enabled = False
Timer2.Enabled = False
TimerFindWindow.Enabled = True
Command1.Caption = "重新开始"
End Sub


Private Sub Command2_Click()
Static i As Boolean
i = Not i
If i = True Then
Command2.Caption = "继续"
Timer1.Enabled = False
 
Else
Command2.Caption = "暂停"
Timer1.Enabled = True
 
End If
'TimerFindWindow.Enabled = True
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
TimerFindWindow.Enabled = False
Timer1.Enabled = False
Timer2.Enabled = False
End Sub


Private Sub Timer1_Timer()

SendMessage dhwnd, WM_KEYDOWN, VK_UP, 0

Timer1.Enabled = False
Timer2.Enabled = True
End Sub

 
Private Sub Timer2_Timer()
PostMessage dhwnd, WM_KEYDOWN, VK_RETURN, 0
PostMessage dhwnd, WM_KEYUP, VK_RETURN, 0
Timer2.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub TimerFindWindow_Timer()
Dim pt  As POINTAPI
Call GetCursorPos(pt)
dhwnd = WindowFromPoint(pt.dx, pt.dy)
Timer1.Enabled = True

TimerFindWindow.Enabled = False

End Sub
