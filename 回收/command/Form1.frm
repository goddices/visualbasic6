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
      Height          =   375
      Index           =   0
      Left            =   -240
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const GAME_WIDTH = 10
Private Const GAME_HEIGHT = 10
'

Private Sub Form_Activate()
Command1(0).Left = -Command1(0).Width
Dim m As Integer
For i = 1 To GAME_WIDTH
    For j = 1 To GAME_HEIGHT
        m = m + 1
        Load Command1(m)
        Command1(m).Caption = m
        Command1(m).Left = i * Command1(m - 1).Left + Command1(m - 1).Width
        
       If m > GAME_HEIGHT Then
            Command1(m).Left = Command1(m - GAME_WIDTH).Left + Command1(m - GAME_WIDTH).Width
            Command1(m).Top = Command1(m - GAME_HEIGHT).Top + Command1(m - GAME_HEIGHT).Height
        End If
        '
      
    
         
        
        Command1(m).Visible = True

    Next
Next
Me.Height = Command1(m).Top + Command1(m).Height + 100
Me.Width = Command1(GAME_WIDTH).Left + Command1(GAME_WIDTH).Width + 100
'For j = 1 To 15
'    Load Command1(j)
'    Command1(j).Top = Command1(j - 1).Top + Command1(j - 1).Height
'    Command1(j).Left = Command1(j - 1).Left
'    Command1(j).Visible = True

End Sub

