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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "st"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   1800
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

Private Const VK_CONTROL = &H11
Private Const KEYEVENTF_KEYUP = &H2

Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
keybd_event 18, 0, 0, 0
keybd_event Asc("Z"), 0, 0, 0
keybd_event Asc("Z"), 0, KEYEVENTF_KEYUP, 0
keybd_event 18, 0, KEYEVENTF_KEYUP, 0
End Sub
