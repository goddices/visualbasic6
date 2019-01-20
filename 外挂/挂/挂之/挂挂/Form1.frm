VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "π“÷Æ"
   ClientHeight    =   1095
   ClientLeft      =   12180
   ClientTop       =   1380
   ClientWidth     =   2115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   2115
   Begin VB.Timer Timer2 
      Left            =   480
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1400
      Left            =   0
      Top             =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
      Call GetHandle
      
End Sub

Private Sub Timer1_Timer()

click
End Sub
