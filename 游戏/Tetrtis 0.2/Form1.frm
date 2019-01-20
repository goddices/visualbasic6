VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tetris  0.3"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   DrawMode        =   1  'Blackness
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   330
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1320
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   144
      TabIndex        =   7
      Top             =   4800
      Width           =   2190
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   3600
      ScaleHeight     =   1110
      ScaleWidth      =   1110
      TabIndex        =   3
      Top             =   720
      Width           =   1140
   End
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   2640
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2640
      Width           =   975
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4400
      Left            =   120
      ScaleHeight     =   292.096
      ScaleMode       =   0  'User
      ScaleWidth      =   219.095
      TabIndex        =   0
      Top             =   360
      Width           =   3300
   End
   Begin VB.Label LblScr 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   180
      Left            =   4440
      TabIndex        =   6
      Top             =   2160
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "得分"
      Height          =   180
      Left            =   3720
      TabIndex        =   5
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "下一个"
      Height          =   180
      Left            =   3720
      TabIndex        =   4
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private i As Integer, j As Integer

Private transform As Byte

Private start As Boolean

Private score As Integer

Private loopTimes As Integer

Private barColor As Long

Private pt  As Point

Private startingBar As Bar

Private theNextBar As Bar

Private theBar As Bar

 
Private Sub Command3_Click()
    Static dv As Boolean
    dv = Not dv
    If dv = True Then
    MsgBox "暂停"
    Timer1.Enabled = False
    Else
    Timer1.Enabled = True
    End If
End Sub

 
Private Sub Form_Activate()

    Call StartGame
End Sub
Private Sub StartGame()
    Pic1.Cls
    dlLenX = Int(Pic1.Width / DLX)
    dlLenY = Int(Pic1.Height / DLY)
    
    pt.ptX = 5
    pt.ptY = -1
    startingBar = RndType()
    theNextBar = RndType()
    start = True
    transform = 3
    Pic1.DrawWidth = 1
    Call TimerInitialize
    
    For i = 0 To 11
        For j = 0 To 15
            Coordinates(i, j) = 0
        Next
    Next
    
    For j = 0 To 16
        Coordinates(-1, j) = 1
        Coordinates(12, j) = 1
    Next
    For i = -1 To 12
        Coordinates(i, -1) = 0
        Coordinates(i, 16) = 1
    Next
'Call AutoArrowDown
End Sub

  
Private Sub Pic1_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim RtIsTrFm As Boolean
    
    Dim RtIsBtm As Boolean
    
    Dim WinLost As Byte
    
    Dim bpt As Point
    
    Pic2.Cls
    
    If start = True Then
        theBar = startingBar
        Call CreateBar(Pic2, bpt, theNextBar.barColor, theNextBar.BarName & CStr(3), True)
    
    End If
    
    Call CreateBar(Pic2, bpt, theNextBar.barColor, theNextBar.BarName & CStr(3), True)
    
    Call CreateBar(Pic1, pt, vbWhite, theBar.BarName & transform)
    RtIsTrFm = IsTransformed(pt, theBar.BarName)
     
    
    Select Case KeyCode
        Case 37
            pt.ptX = pt.ptX - 1
            If IsBound(-1) = True Then
                pt.ptX = pt.ptX + 1
            End If
    
        Case 38
            
            If RtIsTrFm = True Then
                transform = transform + 1
                
                If transform > 4 Then transform = 1
                
            End If
        
        Case 39
            pt.ptX = pt.ptX + 1
            If IsBound(1) = True Then
                pt.ptX = pt.ptX - 1
            End If
            'If (RtPt.ptX >= dlX) Then pt.ptX = pt.ptX - 1
        Case 40
            RtIsBtm = IsBottom()
      
            pt.ptY = pt.ptY + 1
            
            If RtIsBtm = True Then
               
                pt.ptY = pt.ptY - 1
                
                If pt.ptY <= 0 Then
                    MsgBox "you lost"
                    Call StartGame
                    Exit Sub
                End If
                Call BottomProgress
            End If
    End Select
    
    Call CreateBar(Pic1, pt, theBar.barColor, theBar.BarName & transform)
End Sub


Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim a As Integer, b As Integer
        a = Int(x / dlLenX)
        b = Int(y / dlLenY)
     MsgBox Coordinates(a, b) & "  " & a & "  " & b
    
End Sub

  
Private Sub TimerInitialize()
'-_-
    Timer1.Enabled = True
    Timer1.Interval = 1800
End Sub

Private Sub Pic1_Paint()
    Dim dp As Point
    For i = 0 To DLX - 1
        For j = 0 To DLY - 1
            dp.ptX = i
            dp.ptY = j
            If Coordinates(i, j) <> 0 Then BitMap Pic1, dp, Coordinates(i, j)
            
        Next
    Next
End Sub

Private Sub Timer1_Timer()
    Call Pic1_KeyDown(40, 0)
End Sub


Private Sub BottomProgress()
    start = False
    Call CreateBar(Pic1, pt, theBar.barColor, theBar.BarName & transform)
    
    Call IsFullLine
    Call CreateNewBar
    LblScr.Caption = CStr(score)
    loopTimes = 0
    Dim bpt As Point
    Pic2.Cls
    Call CreateBar(Pic2, bpt, theNextBar.barColor, theNextBar.BarName & CStr(3), True)

End Sub


Private Function IsTransformed(LPoint As Point, Optional BarType As String) As Boolean
    Dim Correct As Integer
    IsTransformed = True
    If BarType = "Line" Then
        Correct = 3
    Else
        Correct = 2
    End If
    
    For i = LPoint.ptX To LPoint.ptX + Correct
        For j = LPoint.ptY To LPoint.ptY + Correct
            If (Coordinates(i, j) <> 0) Then IsTransformed = False
        Next
    Next
End Function

Private Function IsBound(MoveStep As Integer) As Boolean
    IsBound = False
    For i = 0 To 7 Step 2
        If Coordinates(Dot(i) + MoveStep, Dot(i + 1)) <> 0 Then IsBound = True
    Next
End Function

Private Function IsBottom() As Boolean
    IsBottom = False
    
    For i = 0 To 7 Step 2
        If Coordinates(Dot(i), Dot(i + 1) + 1) <> 0 Then IsBottom = True
    Next

End Function

Private Sub IsFullLine()
  
    Dim bool  As Boolean
    Dim count As Integer
    Dim ccout    As Integer
    bool = True
    
    Debug.Print pt.ptY
'    For j = 15 To 0 Step -1
    For j = 15 To pt.ptY Step -1
        For i = 0 To 11
            ccout = ccout + 1
            
            bool = bool And CBool(Coordinates(i, j))
        Next
        If bool = True Then
            count = count + 1
            
            KillLines j, count
        End If
        bool = True
    Next
Debug.Print ccout
End Sub


Private Sub KillLines(LineNum As Integer, LinesCount As Integer)
    Dim pt_ As Point
    Dim ColorCode As Byte
    
    loopTimes = loopTimes + LinesCount
    
    score = score + 2 * loopTimes - 1
    'LblScr.Caption = score
    
    pt_.ptX = 0
    pt_.ptY = LineNum
     For i = 0 To DLX - 1
        Coordinates(i, pt_.ptY) = 0
    Next
    
    For i = 0 To DLX - 1
        pt_.ptX = i
        'Call Drawing(Pic1, pt_, vbWhite)
        Call BitMap(Pic1, pt_, 0)
    Next
    
    Call Translation(LineNum, LinesCount)

End Sub

Private Sub Translation(LineNum As Integer, MoveStep As Integer)
    Dim temPoint As Point
    For j = LineNum - MoveStep To 1 Step -1
        For i = 0 To 11
            If (Coordinates(i, j) <> 0 And j <> 0) Then
                Coordinates(i, j + MoveStep) = Coordinates(i, j)
                temPoint.ptX = i
                temPoint.ptY = j
                'Call Drawing(Pic1, temPoint, vbWhite)
                Call BitMap(Pic1, temPoint, 0)
                temPoint.ptY = j + MoveStep
                'Call Drawing(Pic1, temPoint, ColorB2L(Coordinates(i, j)))
                Call BitMap(Pic1, temPoint, Coordinates(i, j))
                Coordinates(i, j) = 0
            End If
        Next
    Next
    Call IsFullLine
End Sub

Private Sub CreateNewBar()

    theBar = theNextBar
    theNextBar = RndType()
    transform = 3
    pt.ptX = 5
    pt.ptY = -1

End Sub

Private Function RndType() As Bar
    Dim num As Integer
    Randomize
    num = Int(Rnd * 7)
    
    Select Case num
        Case 0
            RndType.BarName = "T"
            RndType.barColor = vbGreen
        Case 1
            RndType.BarName = "L"
           RndType.barColor = vbBlue
        Case 2
            RndType.BarName = "CL"
           RndType.barColor = vbMagenta
        Case 3
            RndType.BarName = "Z"
           RndType.barColor = vbRed
        Case 4
            RndType.BarName = "CZ"
           RndType.barColor = vbCyan
        Case 5
            RndType.BarName = "B"
           RndType.barColor = vbGold
        Case 6
            RndType.BarName = "Line"
           RndType.barColor = vbPurple1
    End Select

End Function

