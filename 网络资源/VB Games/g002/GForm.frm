VERSION 5.00
Begin VB.Form GForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stick Men 2"
   ClientHeight    =   5004
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   5940
   Icon            =   "GForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5004
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer MoveT 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1680
      Top             =   120
   End
   Begin VB.PictureBox PB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1332
      Left            =   120
      ScaleHeight     =   111
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   111
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Image PausedL 
      Height          =   492
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Image NMEalertpic 
      Height          =   612
      Index           =   3
      Left            =   1680
      Top             =   2760
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEPic 
      Height          =   612
      Index           =   3
      Left            =   2400
      Top             =   2760
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEangryPic 
      Height          =   612
      Index           =   3
      Left            =   3120
      Top             =   2760
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEconfusedPic 
      Height          =   612
      Index           =   3
      Left            =   3840
      Top             =   2760
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEalertpic 
      Height          =   612
      Index           =   2
      Left            =   1680
      Top             =   2040
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEPic 
      Height          =   612
      Index           =   2
      Left            =   2400
      Top             =   2040
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEangryPic 
      Height          =   612
      Index           =   2
      Left            =   3120
      Top             =   2040
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEconfusedPic 
      Height          =   612
      Index           =   2
      Left            =   3840
      Top             =   2040
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEalertpic 
      Height          =   612
      Index           =   1
      Left            =   1680
      Top             =   1320
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEPic 
      Height          =   612
      Index           =   1
      Left            =   2400
      Top             =   1320
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEangryPic 
      Height          =   612
      Index           =   1
      Left            =   3120
      Top             =   1320
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEconfusedPic 
      Height          =   612
      Index           =   1
      Left            =   3840
      Top             =   1320
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEconfusedPic 
      Height          =   612
      Index           =   0
      Left            =   3840
      Top             =   600
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEangryPic 
      Height          =   612
      Index           =   0
      Left            =   3120
      Top             =   600
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEPic 
      Height          =   612
      Index           =   0
      Left            =   2400
      Top             =   600
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image PowerPic 
      Height          =   612
      Left            =   240
      Top             =   3600
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image BonePic 
      Height          =   612
      Left            =   240
      Top             =   2880
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image NMEalertpic 
      Height          =   612
      Index           =   0
      Left            =   1680
      Top             =   600
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image ManPic 
      Height          =   612
      Left            =   240
      Top             =   2160
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image FloorPic 
      Height          =   612
      Left            =   240
      Top             =   4320
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image WallPic 
      Height          =   612
      Left            =   240
      Top             =   1440
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Image Display 
      BorderStyle     =   1  'Fixed Single
      Height          =   1260
      Left            =   3360
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   2196
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu StartGame 
         Caption         =   "&Start Game"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu LevelDesigner 
         Caption         =   "&Level Designer"
      End
      Begin VB.Menu Line2 
         Caption         =   "-"
      End
      Begin VB.Menu Xit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu Instructions 
         Caption         =   "&Instructions"
      End
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "GForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim picMan(0 To 3, 0 To 3) As IPictureDisp
Dim picBone(0 To 3) As IPictureDisp

Private Sub About_Click()
AboutForm.Visible = True
End Sub

Private Sub File_Click()
MoveT.Enabled = False
PausedL.Visible = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Maze.Object(Man.x, Man.y) = NOWT

Select Case KeyCode
  Case vbKeyP
    If MoveT.Enabled Then
    MoveT.Enabled = False
    PausedL.Visible = True
    Else
    PausedL.Visible = False
    MoveT.Enabled = True
    End If
  Case vbKeyUp
  If MoveT.Enabled = False Then GoTo SkipIt
    If Man.dir = dUP Then
      Select Case Maze.Object(Man.x, Man.y - 1)
        Case NOWT
          Man.y = Man.y - 1
        Case BONE
          sndPlaySound App.Path & "\Resources\Sounds\pop.wav", &H1
          Maze.BoneCount = Maze.BoneCount - 1
          If Maze.BoneCount = 0 Then LevelComplete: Exit Sub
          Man.y = Man.y - 1
          Caption = "Haunted Maze 2 - " & Maze.BoneCount & " bones left to collect"
      End Select
    Else
    Man.dir = dUP
    End If
  Case vbKeyDown
  If MoveT.Enabled = False Then GoTo SkipIt
    If Man.dir = dDOWN Then
      Select Case Maze.Object(Man.x, Man.y + 1)
        Case NOWT
          Man.y = Man.y + 1
        Case BONE
          sndPlaySound App.Path & "\Resources\Sounds\pop.wav", &H1
          Maze.BoneCount = Maze.BoneCount - 1
          If Maze.BoneCount = 0 Then LevelComplete: Exit Sub
          Man.y = Man.y + 1
          Caption = "Haunted Maze 2 - " & Maze.BoneCount & " bones left to collect"
      End Select
    Else
    Man.dir = dDOWN
    End If
  Case vbKeyLeft
  If MoveT.Enabled = False Then GoTo SkipIt
    If Man.dir = dLEFT Then
      Select Case Maze.Object(Man.x - 1, Man.y)
        Case NOWT
          Man.x = Man.x - 1
        Case BONE
          sndPlaySound App.Path & "\Resources\Sounds\pop.wav", &H1
          Maze.BoneCount = Maze.BoneCount - 1
          If Maze.BoneCount = 0 Then LevelComplete: Exit Sub
          Man.x = Man.x - 1
          Caption = "Haunted Maze 2 - " & Maze.BoneCount & " bones left to collect"
      End Select
    Else
    Man.dir = dLEFT
    End If
  Case vbKeyRight
  If MoveT.Enabled = False Then GoTo SkipIt
    If Man.dir = dRIGHT Then
      Select Case Maze.Object(Man.x + 1, Man.y)
        Case NOWT
          Man.x = Man.x + 1
        Case BONE
          sndPlaySound App.Path & "\Resources\Sounds\pop.wav", &H1
          Maze.BoneCount = Maze.BoneCount - 1
          If Maze.BoneCount = 0 Then LevelComplete: Exit Sub
          Man.x = Man.x + 1
          Caption = "Haunted Maze 2 - " & Maze.BoneCount & " bones left to collect"
      End Select
    Else
    Man.dir = dRIGHT
    End If
End Select

SkipIt:
sndPlaySound App.Path & "\Resources\Sounds\step" & Man.Frame & ".wav", &H1
Man.Frame = Man.Frame + 1
If Man.Frame = 4 Then Man.Frame = 0
ManPic = picMan(Man.dir, Man.Frame)
Maze.Object(Man.x, Man.y) = COOLMAN

End Sub

Sub LevelComplete()
MoveT.Enabled = False
sndPlaySound App.Path & "\Resources\Sounds\complete.wav", &H2
Display = LoadPicture()
UnloadMaze
End Sub

Private Sub Form_Load()
Randomize Timer
'lay out the form, then show it
Move 0, 0, Screen.Width, Screen.Height
Display.Move Screen.Width * 0.05, 0, Screen.Width * 0.9, Screen.Height * 0.9
ScaleMode = 3
PB.Move 0, 0, 300, 130
WallPic.Move 0, 0, 32, 32
FloorPic.Move 0, 0, 32, 32
ManPic.Move 0, 0, 32, 32
For i = 0 To 3
NMEangryPic(i).Move 0, 0, 32, 32
NMEPic(i).Move 0, 0, 32, 32
NMEalertpic(i).Move 0, 0, 32, 32
NMEconfusedPic(i).Move 0, 0, 32, 32
Next
ScaleMode = 1
Show
'pre-calculate some co-ords to be used throughout the game
CalcPositions
'put in some default values
Maze.NMEcount = 1
Maze.WallType = 0
Maze.FloorType = 0
ReDim Ghost(0 To Maze.NMEcount)

'load default pics
WallPic = LoadPicture(App.Path & "\Resources\Pictures\Walls\" & Maze.WallType & ".ico")
FloorPic = LoadPicture(App.Path & "\Resources\Pictures\Floors\" & Maze.FloorType & ".ico")

For i = 0 To 3
  NMEangryPic(i) = LoadPicture(App.Path & "\Resources\Pictures\Characters\Ghosts\Angry\" & i & ".ico")
  NMEPic(i) = LoadPicture(App.Path & "\Resources\Pictures\Characters\Ghosts\Normal\" & i & ".ico")
  NMEalertpic(i) = LoadPicture(App.Path & "\Resources\Pictures\Characters\Ghosts\Alert\" & i & ".ico")
  NMEconfusedPic(i) = LoadPicture(App.Path & "\Resources\Pictures\Characters\Ghosts\Confused\" & i & ".ico")
  Set picBone(i) = LoadPicture(App.Path & "\Resources\Pictures\PickUps\Bone" & i + 1 & ".ico")
For i2 = 0 To 3
  Set picMan(i, i2) = LoadPicture(App.Path & "\Resources\Pictures\Characters\CoolMan\" & i & "\" & i2 & ".ico")
Next
Next

ScaleMode = 1
PausedL.Move Screen.Width * 0.3, Screen.Height * 0.3, Screen.Width * 0.4, Screen.Height * 0.4
PausedL = LoadPicture(App.Path & "\Resources\Pictures\Interface\Paused.bmp")
End Sub

Private Sub Form_Unload(Cancel As Integer)
sndPlaySound App.Path & "\Resources\Sounds\exit.wav", &H2
End Sub

Private Sub Help_Click()
MoveT.Enabled = False
PausedL.Visible = True
End Sub

Private Sub Instructions_Click()
HelpForm.Visible = True
End Sub

Private Sub LevelDesigner_Click()
DForm.Visible = True
Unload Me
End Sub

Private Sub MoveT_Timer()
On Error Resume Next
Maze.Frame = Maze.Frame + 1
If Maze.Frame = 4 Then Maze.Frame = 0
BonePic = picBone(Maze.Frame)
For i = 0 To Maze.NMEcount
Select Case Ghost(i).State
  Case NORMAL: TempByte = 5
  Case ANGRY: TempByte = 3
  Case ALERT: TempByte = 2
  Case CONFUSED: TempByte = 7
End Select
If CanSee Then
MoveGhost dFOWARD, False
Else
    Select Case Int(Rnd * TempByte)
      Case 0
        Select Case Man.x
        Case Is < Ghost(i).x: MoveGhost dLEFT, True
        Case Is > Ghost(i).x: MoveGhost dRIGHT, True
        End Select
      Case 1
        Select Case Man.y
        Case Is < Ghost(i).y: MoveGhost dUP, True
        Case Is > Ghost(i).y: MoveGhost dDOWN, True
        End Select
    End Select
End If
Next
PB.Cls 'clear backbuffer
For i = 1 To 14 'loop through the camera view,
For i2 = 1 To 15 'drawing each object
  Select Case Maze.Object(i + Man.x - 8, i2 + Man.y - 7)
  Case WALL
  PB.PaintPicture WallPic, Position(i, i2).x, Position(i, i2).y
  Case COOLMAN
  PB.PaintPicture ManPic, Position(i, i2).x, Position(i, i2).y
  Case BONE
  PB.PaintPicture BonePic, Position(i, i2).x, Position(i, i2).y
  Case NME0
  PB.PaintPicture NMEPic(0), Position(i, i2).x, Position(i, i2).y
  Case NME1
  PB.PaintPicture NMEPic(1), Position(i, i2).x, Position(i, i2).y
  Case NME2
  PB.PaintPicture NMEPic(2), Position(i, i2).x, Position(i, i2).y
  Case NME3
  PB.PaintPicture NMEPic(3), Position(i, i2).x, Position(i, i2).y
  Case NMEalert0
  PB.PaintPicture NMEalertpic(0), Position(i, i2).x, Position(i, i2).y
  Case NMEalert1
  PB.PaintPicture NMEalertpic(1), Position(i, i2).x, Position(i, i2).y
  Case NMEalert2
  PB.PaintPicture NMEalertpic(2), Position(i, i2).x, Position(i, i2).y
  Case NMEalert3
  PB.PaintPicture NMEalertpic(3), Position(i, i2).x, Position(i, i2).y
  Case NMEangry0
  PB.PaintPicture NMEangryPic(0), Position(i, i2).x, Position(i, i2).y
  Case NMEangry1
  PB.PaintPicture NMEangryPic(1), Position(i, i2).x, Position(i, i2).y
  Case NMEangry2
  PB.PaintPicture NMEangryPic(2), Position(i, i2).x, Position(i, i2).y
  Case NMEangry3
  PB.PaintPicture NMEangryPic(3), Position(i, i2).x, Position(i, i2).y
  Case NMEconfused0
  PB.PaintPicture NMEconfusedPic(0), Position(i, i2).x, Position(i, i2).y
  Case NMEconfused1
  PB.PaintPicture NMEconfusedPic(1), Position(i, i2).x, Position(i, i2).y
  Case NMEconfused2
  PB.PaintPicture NMEconfusedPic(2), Position(i, i2).x, Position(i, i2).y
  Case NMEconfused3
  PB.PaintPicture NMEconfusedPic(3), Position(i, i2).x, Position(i, i2).y
  End Select
Next
Next
Display = PB.Image 'copy to final display
End Sub

Function CanSee() As Boolean
On Error GoTo OuttaHere
Select Case Ghost(i).dir
  Case dUP
  For Temp.y = 1 To 6
  Select Case Maze.Object(Ghost(i).x, Ghost(i).y - Temp.y)
     Case COOLMAN: CanSee = True: Exit Function
     Case WALL: Exit Function
  End Select
  Next
  Case dDOWN
  For Temp.y = 1 To 6
  Select Case Maze.Object(Ghost(i).x, Ghost(i).y + Temp.y)
     Case COOLMAN: CanSee = True: Exit Function
     Case WALL: Exit Function
  End Select
  Next
  Case dLEFT
  For Temp.x = 1 To 6
  Select Case Maze.Object(Ghost(i).x - Temp.x, Ghost(i).y)
     Case COOLMAN: CanSee = True: Exit Function
     Case WALL: Exit Function
  End Select
  Next
  Case dRIGHT
  For Temp.x = 1 To 6
  Select Case Maze.Object(Ghost(i).x + Temp.x, Ghost(i).y)
     Case COOLMAN: CanSee = True: Exit Function
     Case WALL: Exit Function
  End Select
  Next
End Select
OuttaHere:
End Function

Private Sub StartGame_Click()
Lives = 3
SelectLevelForm.Visible = True
End Sub


Private Sub Xit_Click()
Unload Me
End Sub

Sub PlaceAllObjects()

For i = 0 To Maze.Size
For i2 = 0 To Maze.Size
  Select Case Maze.Object(i, i2)
    Case WALL
    Case NOWT
    Case Else
    Maze.Object(i, i2) = NOWT
  End Select
Next
Next

Maze.NMEcount = Maze.Size \ 6
Maze.BoneCount = Maze.Size \ 3
ReDim Ghost(0 To Maze.NMEcount)
'select random positions for each object
For i = 1 To Maze.BoneCount
Retry:
  tempx = Int(Rnd * Maze.Size)
  tempy = Int(Rnd * Maze.Size)
  If Maze.Object(tempx, tempy) = NOWT Then
     Maze.Object(tempx, tempy) = BONE
  Else 'if something is in that position,
  GoTo Retry 'then try another position
  End If
Next
For i = 0 To Maze.NMEcount
Retry3:
  tempx = Int(Rnd * Maze.Size)
  tempy = Int(Rnd * Maze.Size)
  If Maze.Object(tempx, tempy) = NOWT Then
  Else 'if something is in that position,
  GoTo Retry3 'then try another position
  End If
  Ghost(i).x = tempx
  Ghost(i).y = tempy
  Ghost(i).dir = Int(Rnd * 4)
  Select Case Ghost(i).dir
    Case dUP
    Maze.Object(Ghost(i).x, Ghost(i).y) = NME0
    Case dRIGHT
    Maze.Object(Ghost(i).x, Ghost(i).y) = NME1
    Case dDOWN
    Maze.Object(Ghost(i).x, Ghost(i).y) = NME2
    Case dLEFT
    Maze.Object(Ghost(i).x, Ghost(i).y) = NME3
  End Select
Next

Retry2:
  tempx = Int(Rnd * Maze.Size)
  tempy = Int(Rnd * Maze.Size)
  If Maze.Object(tempx, tempy) = NOWT Then
     Maze.Object(tempx, tempy) = COOLMAN
  Else 'if something is in that position,
  GoTo Retry2 'then try another position
  End If
  Man.x = tempx
  Man.y = tempy
  Man.dir = Int(Rnd * 4)
  Maze.Object(Man.x, Man.y) = COOLMAN
End Sub

Sub PaintFloor()
WallPic = LoadPicture(App.Path & "\Resources\Pictures\Walls\" & Maze.WallType & ".ico")
FloorPic = LoadPicture(App.Path & "\Resources\Pictures\Floors\" & Maze.FloorType & ".ico")
PB.Cls 'clear pic
For i = -2 To 20
For i2 = -3 To 20
  PB.PaintPicture FloorPic, (i * 24) - (i2 * 8) - 10, i2 * 8
Next
Next
PB.Picture = PB.Image 'store the image in pic property
End Sub

Sub MoveGhost(dir As Byte, Retry As Boolean)
i5 = 0
On Error Resume Next
Maze.Object(Ghost(i).x, Ghost(i).y) = NOWT

If dir = dFOWARD Then
  Ghost(i).State = ALERT
    Select Case Ghost(i).dir
    Case dUP: Ghost(i).y = Ghost(i).y - 1
    Case dDOWN: Ghost(i).y = Ghost(i).y + 1
    Case dLEFT: Ghost(i).x = Ghost(i).x - 1
    Case dRIGHT: Ghost(i).x = Ghost(i).x + 1
    End Select
  Temp.x = 0: Temp.y = 0
Else
TryAgain:
i5 = i5 + 1
Select Case i5
  Case 1: Ghost(i).State = ANGRY
  'Case 2: Ghost(i).State = NORMAL
  Case 2: Ghost(i).State = CONFUSED
  Case 4: Retry = False
End Select
Ghost(i).dir = dir
Select Case dir
  Case dUP: Temp.x = 0: Temp.y = -1
  Case dDOWN: Temp.x = 0: Temp.y = 1
  Case dLEFT: Temp.x = -1: Temp.y = 0
  Case dRIGHT: Temp.x = 1: Temp.y = 0
End Select
End If

Select Case Maze.Object(Ghost(i).x + Temp.x, Ghost(i).y + Temp.y)
  Case NOWT
    Maze.Object(Ghost(i).x, Ghost(i).y) = NOWT
    Ghost(i).x = Ghost(i).x + Temp.x
    Ghost(i).y = Ghost(i).y + Temp.y
  Case COOLMAN
    'PaintLevel
    sndPlaySound App.Path & "\Resources\Sounds\dead.wav", &H2
    PlaceAllObjects
  Case Else
    If Retry Then
      dir = Int(Rnd * 4)
      GoTo TryAgain
    End If
End Select

Select Case Ghost(i).State
Case NORMAL
  Select Case Ghost(i).dir
    Case dUP
    Maze.Object(Ghost(i).x, Ghost(i).y) = NME0
    Case dRIGHT
    Maze.Object(Ghost(i).x, Ghost(i).y) = NME1
    Case dDOWN
    Maze.Object(Ghost(i).x, Ghost(i).y) = NME2
    Case dLEFT
    Maze.Object(Ghost(i).x, Ghost(i).y) = NME3
  End Select
Case ANGRY
  Select Case Ghost(i).dir
    Case dUP
    Maze.Object(Ghost(i).x, Ghost(i).y) = NMEangry0
    Case dRIGHT
    Maze.Object(Ghost(i).x, Ghost(i).y) = NMEangry1
    Case dDOWN
    Maze.Object(Ghost(i).x, Ghost(i).y) = NMEangry2
    Case dLEFT
    Maze.Object(Ghost(i).x, Ghost(i).y) = NMEangry3
  End Select
Case ALERT
  Select Case Ghost(i).dir
    Case dUP
    Maze.Object(Ghost(i).x, Ghost(i).y) = NMEalert0
    Case dRIGHT
    Maze.Object(Ghost(i).x, Ghost(i).y) = NMEalert1
    Case dDOWN
    Maze.Object(Ghost(i).x, Ghost(i).y) = NMEalert2
    Case dLEFT
    Maze.Object(Ghost(i).x, Ghost(i).y) = NMEalert3
  End Select
Case CONFUSED
  Select Case Ghost(i).dir
    Case dUP
    Maze.Object(Ghost(i).x, Ghost(i).y) = NMEconfused0
    Case dRIGHT
    Maze.Object(Ghost(i).x, Ghost(i).y) = NMEconfused1
    Case dDOWN
    Maze.Object(Ghost(i).x, Ghost(i).y) = NMEconfused2
    Case dLEFT
    Maze.Object(Ghost(i).x, Ghost(i).y) = NMEconfused3
  End Select
End Select
End Sub

