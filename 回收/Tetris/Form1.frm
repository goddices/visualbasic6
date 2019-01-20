VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "for testing"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8850
   DrawMode        =   1  'Blackness
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   611
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   590
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2100
      Left            =   6600
      ScaleHeight     =   2070
      ScaleWidth      =   2070
      TabIndex        =   3
      Top             =   1200
      Width           =   2100
   End
   Begin VB.Timer Timer1 
      Left            =   8160
      Top             =   5040
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   5640
      Width           =   975
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7995
      Left            =   480
      ScaleHeight     =   533
      ScaleMode       =   0  'User
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   480
      Width           =   6000
   End
   Begin VB.Label LblScr 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7680
      TabIndex        =   6
      Top             =   3720
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "得分"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6720
      TabIndex        =   5
      Top             =   3720
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "下一个"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6600
      TabIndex        =   4
      Top             =   480
      Width           =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private i As Integer, j As Integer

Private Wd As Integer
Private Ht As Integer

Private dlWd As Single
Private dlHt As Single

Private Transform As Byte

Private Start As Boolean

Private Score As Integer

Private LoopTimes As Integer

Private BarColor As Long

Private pt  As Point

Private StartingBar As Bar

Private theNextBar As Bar

Private theBar As Bar

Private Sub Command1_Click()
Pic1.DrawWidth = 1
For i = 5 To 6
    For j = 5 To 6
        Pic1.Line (i * dlWd, 0)-(i * dlWd, i * dlHt), vbRed
        Pic1.Line (0, i * dlHt)-(i * dlWd, i * dlHt), vbRed
    Next
Next
End Sub

 
 

Private Sub Command2_Click()

End Sub

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

Private Sub Command5_Click()

End Sub

Private Sub Form_Activate()
 
Wd = Pic1.Width
Ht = Pic1.Height

dlWd = Wd / dlX
dlHt = Ht / dlY

dlLenX = Int(dlWd)
dlLenY = Int(dlHt)

pt.ptX = 5
pt.ptY = -1
Call StartGame

End Sub
Private Sub StartGame()
Pic1.Cls
StartingBar = RndType()
theNextBar = RndType()
Start = True
Transform = 3
Pic1.DrawWidth = 1
Call TimerInitialize
'For i = 1 To dlX
   '  Pic1.Line (i * dlWd, 0)-(i * dlWd, Ht)
'Next

'For i = 1 To dlY
'    Pic1.Line (0, i * dlHt)-(Wd, i * dlHt)
'Next


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

If Start = True Then
    theBar = StartingBar
    Call CreateBar(Pic2, bpt, theNextBar.BarColor, theNextBar.BarName & CStr(3), "NextBar")

End If

Call CreateBar(Pic2, bpt, theNextBar.BarColor, theNextBar.BarName & CStr(3), "NextBar")

Call CreateBar(Pic1, pt, vbWhite, theBar.BarName & Transform)
RtIsTrFm = IsTransformed(pt, theBar.BarName)
 

Select Case KeyCode
    Case 37
        pt.ptX = pt.ptX - 1
        If IsBound(-1) = True Then
            pt.ptX = pt.ptX + 1
        End If

    Case 38
        
        If RtIsTrFm = True Then
            Transform = Transform + 1
            
            If Transform > 4 Then Transform = 1
            
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
            
            If pt.ptY <= 0 Then MsgBox "you lost": Call StartGame
            Call BottomProgress
        End If
End Select

Call CreateBar(Pic1, pt, theBar.BarColor, theBar.BarName & Transform)
End Sub

 
Private Sub BottomProgress()
Start = False
Call CreateBar(Pic1, pt, theBar.BarColor, theBar.BarName & Transform)
LblScr.Caption = CStr(Score)
Call IsFullLine
Call CreateNewBar
LoopTimes = 0
Dim bpt As Point
Pic2.Cls
Call CreateBar(Pic2, bpt, theNextBar.BarColor, theNextBar.BarName & CStr(3), "NextBar")

End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As Integer, b As Integer
    a = Int(X / dlLenX)
    b = Int(Y / dlLenY)
 MsgBox Coordinates(a, b) & "  " & a & "  " & b
    
End Sub

Private Sub AutoArrowDown()
Call Pic1_KeyDown(40, 0)
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
'IsFullLine = False
'Dim IsFullLine_ As Boolean

Dim bool  As Boolean
Dim count As Integer
bool = True
For j = 15 To 0 Step -1
    For i = 0 To 11
        bool = bool And CBool(Coordinates(i, j))
    Next
    If bool = True Then
        count = count + 1
        Print j; count
        KillLines j, count
    End If
    bool = True
Next

End Sub


Private Sub KillLines(LineNum As Integer, LinesCount As Integer)
Dim pt_ As Point
LoopTimes = LoopTimes + LinesCount

Score = Score + 2 * LoopTimes - 1
LblScr.Caption = Score

pt_.ptX = 0
pt_.ptY = LineNum
'Print LineNum & " " & LinesCount
Call CreateBar(Pic1, pt_, vbWhite, "KillLines", "KillLines")
Call Translation(LineNum, LinesCount)

End Sub

Private Sub Translation(LineNum As Integer, MoveStep As Integer)
Dim apt As Point

For j = LineNum - MoveStep To 1 Step -1
    For i = 0 To 11
        If (Coordinates(i, j) <> 0 And j <> 0) Then
            Coordinates(i, j + MoveStep) = Coordinates(i, j)
            apt.ptX = i
            apt.ptY = j
            Call CreateBar(Pic1, apt, vbWhite, "Translate", "Translate")
            
            apt.ptY = j + MoveStep
            Call CreateBar(Pic1, apt, ColorB2L(Coordinates(i, j)), "Translate", "Translate")
            Coordinates(i, j) = 0
        End If
    Next
Next
Call IsFullLine
End Sub

Private Sub CreateNewBar()

theBar = theNextBar
theNextBar = RndType()
Transform = 3
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
        RndType.BarColor = vbGreen
    Case 1
        RndType.BarName = "L"
       RndType.BarColor = vbBlue
    Case 2
        RndType.BarName = "CL"
       RndType.BarColor = vbMagenta
    Case 3
        RndType.BarName = "Z"
       RndType.BarColor = vbRed
    Case 4
        RndType.BarName = "CZ"
       RndType.BarColor = vbCyan
    Case 5
        RndType.BarName = "B"
       RndType.BarColor = vbGold
    Case 6
        RndType.BarName = "Line"
       RndType.BarColor = vbPurple1
End Select

End Function

 
Private Sub TimerInitialize()
'-_-
Timer1.Enabled = True
Timer1.Interval = 1800
End Sub



Private Sub Timer1_Timer()
Call AutoArrowDown
End Sub
