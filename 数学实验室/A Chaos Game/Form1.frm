VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   499
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   525
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "1000"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "execute"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "iterative times"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINT
    x As Double
    y As Double
End Type

Private A As POINT
Private B As POINT
Private C As POINT

Private Function MidPoint(pt1 As POINT, pt2 As POINT) As POINT
    MidPoint.x = (pt1.x + pt2.x) / 2
    MidPoint.y = (pt1.y + pt2.y) / 2
End Function


Private Sub Give(pt As POINT, ByVal x As Double, ByVal y As Double)
    '*********************
    'Assignment
    pt.x = x
    pt.y = y
End Sub

Private Sub Command1_Click()
    Cls
    Triangle
    Dim times As Long
    Dim pt  As POINT
    Dim D As Integer ' Dices

    If IsNumeric(Text1.Text) Then
        
        times = CLng(Text1.Text)
        Give pt, Rnd * 100, Rnd * 100
        
        Randomize
        
        For i = 1 To times
        
            D = Int(Rnd * 3) + 1
            If D = 1 Then
                pt = MidPoint(pt, A)
            ElseIf D = 2 Then
                pt = MidPoint(pt, B)
            Else
                pt = MidPoint(pt, C)
            End If
            
            PSet (pt.x, pt.y)
        Next
    End If
End Sub

Private Sub Form_Activate()
'p1(0,80)
'p2(-80,-80)
'p3(80,-80)
 Triangle
End Sub

Private Sub Form_Load()
Give A, 0, 80
Give B, -80, -80
Give C, 80, -80
Me.Scale (-100, 100)-(100, -100)

End Sub

 
Private Sub Triangle()
    Me.Line (0, 80)-(-80, -80)
    Me.Line (0, 80)-(80, -80)
    Me.Line (-80, -80)-(80, -80)

End Sub
