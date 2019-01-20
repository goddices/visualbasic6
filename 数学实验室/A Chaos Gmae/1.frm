VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "混沌游戏"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   14340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   565
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   956
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   720
      TabIndex        =   1
      Text            =   "4"
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2400
      TabIndex        =   0
      Text            =   "140"
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "循环次数"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "n边形"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINT
    X As Double
    Y As Double
End Type

Private Const MAX_NUM As Integer = 10
Private Const MIN_NUM As Integer = 3

Private pts() As POINT
Private ptsnum As Integer
Dim times As Long

Private Function MidPoint(pt1 As POINT, pt2 As POINT) As POINT
    MidPoint.X = (pt1.X + pt2.X) / 2
    MidPoint.Y = (pt1.Y + pt2.Y) / 2
End Function


Private Sub Give(pt As POINT, ByVal X As Double, ByVal Y As Double)
    '*********************
    'Assignment
    pt.X = X
    pt.Y = Y
End Sub

Private Sub Form_Activate()

    Me.BackColor = vbBlack
    Me.ScaleMode = 3
     Me.Width = 7000
     Me.Height = 7000
    Me.Scale (-100, -100)-(100, 100)
  
    Text2.SelLength = Len(Text2.Text)
End Sub

Private Sub DrawShape(shapeType As Integer)
    
    Select Case shapeType
        Case 3 ' triangle
            Give pts(0), 0, -80
            Give pts(1), -80, 80
            Give pts(2), 80, 80
            DrawConnectLine pts(0), pts(1)
            DrawConnectLine pts(1), pts(2)
             DrawConnectLine pts(2), pts(0)
           
        Case 4
            Give pts(0), -80, -80 ' left top
            Give pts(1), -80, 80  'left bottom
            Give pts(2), 80, 80  'right bottom
            Give pts(3), 80, -80  'right top
            
            DrawConnectLine pts(0), pts(1)
            DrawConnectLine pts(1), pts(2)
            DrawConnectLine pts(2), pts(3)
            DrawConnectLine pts(3), pts(0)
        Case 5
            Give pts(0), 0, -80 '
            Give pts(1), -80, 80  '
            Give pts(2), 80, 80  '
            Give pts(3), 80, -80  '
            Give pts(3), 80, -80  '
            
            DrawConnectLine pts(0), pts(1)
            DrawConnectLine pts(1), pts(2)
            DrawConnectLine pts(2), pts(3)
            DrawConnectLine pts(3), pts(0)
        Case 6
        
    End Select

End Sub


Private Sub DrawConnectLine(pt1 As POINT, pt2 As POINT)
    Me.Line (pt1.X, pt1.Y)-(pt2.X, pt2.Y), vbWhite
             
End Sub


Private Sub Draw()
    Cls
    
    Dim pt  As POINT
    Dim Dice As Integer ' Dices

    Call DrawShape(ptsnum)
   
    Give pt, Rnd * 100, Rnd * 100
    Randomize
    DoEvents
    For i = 1 To times
        Dice = Int(Rnd * ptsnum)
        pt = MidPoint(pt, pts(Dice))
        PSet (pt.X, pt.Y), vbWhite
    Next
     
End Sub

 
Private Sub Text1_KeyDown(Keycode As Integer, Shift As Integer)
     CheckTexts (Keycode)
End Sub

Private Sub Text2_KeyDown(Keycode As Integer, Shift As Integer)
    CheckTexts (Keycode)
End Sub

Private Sub CheckTexts(Keycode As Integer)
    If Keycode = 13 And IsNumeric(Text1.Text) And IsNumeric(Text2.Text) Then
        ptsnum = CLng(Text1.Text)
        If ptsnum > MAX_NUM Or ptsnum < MIN_NUM Then ptsnum = 3
        ReDim pts(ptsnum - 1) As POINT
        times = CLng(Text2.Text)
        Call Draw
    End If
End Sub
