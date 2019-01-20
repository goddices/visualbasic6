VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "麦克石膏飞"
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   3375
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2400
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1440
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "结束"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "继续"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "暂停"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "开始"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long _
)
Private Const VK_RETURN = &HD
Private Const VK_UP = &H26

Private Sub Command1_Click()
End
End Sub



Private Sub Command2_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
End Sub
 

Private Sub Command3_Click()
Timer1.Enabled = True
Timer2.Enabled = True
Command3.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Command4_Click()
Timer1.Enabled = True
Timer2.Enabled = True
Command2.Enabled = True
 End Sub

 

Private Sub Form_Load()
Timer1.Enabled = False
Timer2.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
 
End Sub

Private Sub Timer1_Timer()
Call keybd_event(VK_RETURN, 0, 0, 0)
Timer1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Call keybd_event(VK_UP, 0, 0, 0)
Timer2.Enabled = False
Timer1.Enabled = True
End Sub
