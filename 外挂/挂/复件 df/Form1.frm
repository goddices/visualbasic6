VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "麦克石膏飞"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6075
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   5760
      Top             =   0
   End
   Begin VB.CommandButton CmdReset 
      Caption         =   "重设"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CmdCnfm 
      Caption         =   "确定"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "设置时间（分钟）"
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "结束"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5160
      Top             =   0
   End
   Begin VB.Timer TimerFindWindow 
      Interval        =   5000
      Left            =   4320
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "暂停"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1905
      Left            =   120
      Picture         =   "Form1.frx":0000
      Top             =   120
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "最大20分钟，最小1分钟"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
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

Private Sub CmdCnfm_Click()
If IsNumeric(Text1.Text) Then
    MaxTime = CInt(Text1.Text)
    If MaxTime > 20 Then
        MsgBox "大于20分钟", vbOKOnly, "麦考石膏飞"
        Text1.SetFocus
        Text1.SelLength = Len(Text1.Text)
        Exit Sub
    End If
    MsgBox "你已设置 " & MaxTime & "分钟", vbOKOnly + vbInformation, "麦考石膏飞"
Else
    MsgBox "不允许非数字", vbOKOnly, "麦考石膏飞"
    Text1.SetFocus
    Text1.SelLength = Len(Text1.Text)
    Exit Sub
End If


Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
CmdReset.Enabled = True
CmdCnfm.Enabled = False
End Sub

Private Sub CmdReset_Click()
Timer1.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
CmdCnfm.Enabled = True
CmdReset.Enabled = False
End Sub

Private Sub Command1_Click()
mm_execute
Timer1.Enabled = True
 
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
 
' Command1.Enabled = False
' Command2.Enabled = False
 'Command3.Enabled = False
 
' CmdReset.Enabled = False
 
 'Text1.SelLength = Len(Text1.Text)
End Sub


Private Sub Timer1_Timer()
PostMessage dhwnd, WM_CHAR, 97, 0
PostMessage dhwnd, WM_KEYDOWN, VK_RETURN, 0 '&H20000001 'ALT键按下
PostMessage dhwnd, WM_KEYUP, VK_RETURN, 0 '&H20000001 'E键按下必须要把第29位设置成1，代表alt键已经下
'PostMessage dhwnd, WM_SYSCHAR, 83, &H20000001 ' 发送一个系统字符E
'PostMessage dhwnd, WM_SYSKEYUP, 83, &H80000001 'E键放开，必须把31位设置成1，表示这个是系统键
'PostMessage dhwnd, WM_KEYUP, VK_MENU, &H80000001 ' ALT键放开，31位系统键设置成1
'Timer1.Enabled = False
'Timer2.Enabled = True
'mm_execute
'If Ttime() Then mm_execute
End Sub

Private Sub Timer2_Timer()
PostMessage dhwnd, WM_CHAR, 97, 0
PostMessage dhwnd, WM_SYSKEYDOWN, VK_MENU, &H20000001 'ALT键按下
PostMessage dhwnd, WM_SYSKEYDOWN, 83, &H20000001 'E键按下必须要把第29位设置成1，代表alt键已经下
PostMessage dhwnd, WM_SYSCHAR, 83, &H20000001  ' 发送一个系统字符E
PostMessage dhwnd, WM_SYSKEYUP, 83, &H80000001  'E键放开，必须把31位设置成1，表示这个是系统键
PostMessage dhwnd, WM_KEYUP, VK_MENU, &H80000001 ' ALT键放开，31位系统键设置成1
Timer2.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub mm_execute()
PostMessage dhwnd, WM_SYSKEYDOWN, VK_MENU, &H20000001 'ALT键按下
PostMessage dhwnd, WM_SYSKEYDOWN, 90, &H20000001 'E键按下必须要把第29位设置成1，代表alt键已经下
PostMessage dhwnd, WM_SYSCHAR, 90, &H20000001 ' 发送一个系统字符E
PostMessage dhwnd, WM_SYSKEYUP, 90, &H80000001 'E键放开，必须把31位设置成1，表示这个是系统键
PostMessage dhwnd, WM_KEYUP, VK_MENU, &H80000001 ' ALT键放开，31位系统键设置成1


PostMessage dhwnd, WM_SYSKEYDOWN, VK_MENU, &H20000001 'ALT键按下
PostMessage dhwnd, WM_SYSKEYDOWN, 90, &H20000001 'E键按下必须要把第29位设置成1，代表alt键已经下
PostMessage dhwnd, WM_SYSCHAR, 90, &H20000001 ' 发送一个系统字符E
PostMessage dhwnd, WM_SYSKEYUP, 90, &H80000001 'E键放开，必须把31位设置成1，表示这个是系统键
PostMessage dhwnd, WM_KEYUP, VK_MENU, &H80000001 ' ALT键放开，31位系统键设置成1

End Sub
Private Sub TimerFindWindow_Timer()
Dim pt  As POINTAPI
Call GetCursorPos(pt)
dhwnd = WindowFromPoint(pt.dx, pt.dy)
Timer1.Enabled = True

TimerFindWindow.Enabled = False

End Sub
