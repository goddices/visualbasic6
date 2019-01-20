VERSION 5.00
Object = "{13F8219D-9E97-45E0-92ED-3A225E1833A4}#13.0#0"; "wocao.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6960
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin RunningTime.TimerBar TimerBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1200
      Top             =   2280
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Randomize
TimerBar1.BarColor = vbGreen
TimerBar1.TimerSwitch = True
End Sub

Private Sub Command2_Click()
TimerBar1.TimerSwitch = False
End Sub

Private Sub Command3_Click()
 
Print TimerBar1.IsTimeUp
End Sub

Private Sub Command4_Click()
TimerBar1.ReStart
 
End Sub

 
Private Sub Command6_Click()
TimerBar1.QuickStart True, RGB(123, 123, 123), 10
 
End Sub

Private Sub Timer1_Timer()
Dim i As Long
i = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
'TimerBar1.BarColor = i
End Sub
