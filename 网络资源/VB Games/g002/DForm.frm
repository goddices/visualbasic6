VERSION 5.00
Begin VB.Form DForm 
   Caption         =   "Stick Men 2 Maze Designer"
   ClientHeight    =   4224
   ClientLeft      =   132
   ClientTop       =   372
   ClientWidth     =   5964
   Icon            =   "DForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FloorF 
      Caption         =   "Floor Type"
      Height          =   2052
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1812
      Begin VB.CommandButton cmdChangeFloor 
         Caption         =   "Change Floor Type"
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1572
      End
      Begin VB.Image FloorPic 
         BorderStyle     =   1  'Fixed Single
         Height          =   1332
         Left            =   240
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1332
      End
   End
   Begin VB.Frame WallF 
      Caption         =   "Wall Type"
      Height          =   2052
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1812
      Begin VB.CommandButton cmdChangeWall 
         Caption         =   "Change Wall Type"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1572
      End
      Begin VB.Image WallPic 
         BorderStyle     =   1  'Fixed Single
         Height          =   1332
         Left            =   240
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1332
      End
   End
   Begin VB.PictureBox VPB 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3852
      Left            =   2040
      ScaleHeight     =   3852
      ScaleWidth      =   3852
      TabIndex        =   1
      Top             =   0
      Width           =   3852
      Begin VB.Shape Box 
         BorderColor     =   &H000000FF&
         Height          =   492
         Left            =   360
         Top             =   240
         Width           =   492
      End
   End
   Begin VB.PictureBox PB 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1212
      Left            =   -1200
      ScaleHeight     =   1212
      ScaleWidth      =   1332
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label CoOrdL 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Co-Ords : 0,0"
      Height          =   312
      Left            =   2040
      TabIndex        =   4
      Top             =   3840
      Width           =   3852
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu NewMaze 
         Caption         =   "&New Level"
      End
      Begin VB.Menu SaveMaze 
         Caption         =   "&Save Level"
      End
      Begin VB.Menu LoadMaze 
         Caption         =   "&Load Level"
      End
      Begin VB.Menu DeleteLevel 
         Caption         =   "&Delete Level"
      End
      Begin VB.Menu Xit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "DForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cursor As TCoOrd
Dim LastMouse As TCoOrd
Dim xx As Byte
Dim yy As Byte

Public Sub CreateBorderedMaze(Size As Byte, WallType As Byte)
MoveCursor 0, 0
UnloadMaze
CreateMaze Size, WallType 'create a new maze of the size and style given to the procedure
For i = 0 To Size 'this makes the outer walls
    Maze.Object(0, i) = WALL
    Maze.Object(i, 0) = WALL
    Maze.Object(Size - 1, i) = WALL
    Maze.Object(i, Size - 1) = WALL
Next
DrawRoughMaze
End Sub

Private Sub cmdChangeFloor_Click()
'change to next type of floor
Maze.FloorType = Maze.FloorType + 1
If Maze.FloorType = NUMFLOORS + 1 Then Maze.FloorType = 0
'load the new picture
FloorPic.Picture = LoadPicture(App.Path & "\Resources\Pictures\Floors\" & Maze.FloorType & ".ico")
End Sub

Private Sub cmdChangeWall_Click()
'change to next type of wall
Maze.WallType = Maze.WallType + 1
If Maze.WallType = NUMWALLS + 1 Then Maze.WallType = 0
'load the new picture
WallPic.Picture = LoadPicture(App.Path & "\Resources\Pictures\Walls\" & Maze.WallType & ".ico")
End Sub

Private Sub DeleteLevel_Click()
SaveLevelForm.Caption = "Delete Level"
SaveLevelForm.FileNameT.Visible = False
SaveLevelForm.cmdSave.Caption = "Delete"
SaveLevelForm.Visible = True
End Sub

Private Sub Form_Load()
'create a default bordered maze
CreateBorderedMaze 20, 0 'create a default size and style maze
Show
Tool = PENCIL 'set drawing tool
DrawRoughMaze 'draw out the new maze
'now load some default pics
WallPic = LoadPicture(App.Path & "\Resources\Pictures\Walls\" & Maze.WallType & ".ico")
FloorPic = LoadPicture(App.Path & "\Resources\Pictures\Floors\" & Maze.FloorType & ".ico")
Maze.WallType = 0
Maze.FloorType = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
GForm.Visible = True
End Sub

Private Sub LoadMaze_Click()
SaveLevelForm.Caption = "Load Level"
SaveLevelForm.FileNameT.Visible = False
SaveLevelForm.cmdSave.Caption = "Load"
SaveLevelForm.Visible = True
End Sub

Private Sub NewMaze_Click()
NewLevelForm.Visible = True
End Sub

Private Sub SaveMaze_Click()
SaveLevelForm.Caption = "Save Level"
SaveLevelForm.FileNameT.Visible = True
SaveLevelForm.cmdSave.Caption = "Save"
SaveLevelForm.Visible = True
End Sub

Private Sub VPB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
x = Int(x)
y = Int(y)
xx = x
yy = y
LastMouse.x = x 'move cursor and record mouse co-ords
LastMouse.y = y
MoveCursor xx, yy
End Sub

Private Sub VPB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
CoOrdL = "Co-Ords : " & Int(x) & "," & Int(y)
End Sub

Private Sub VPB_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
x = Int(x)
y = Int(y)
xx = x
yy = y
MoveCursor xx, yy 'move cursor
PB.Cls

'draw a wall
PB.Line (LastMouse.x, LastMouse.y)-(x, y), vbRed, BF
PB.PSet (x, y), vbRed

For i = 0 To Maze.Size
For i2 = 0 To Maze.Size
  If PB.Point(i, i2) = vbRed Then
      If Maze.Object(i, i2) = WALL Then
      Maze.Object(i, i2) = NOWT
      Else
      Maze.Object(i, i2) = WALL
      End If
  End If
Next
Next

DrawRoughMaze 'update the pic of maze
End Sub

Private Sub VPB_Paint()
DrawRoughMaze
End Sub

Private Sub Xit_Click()
Unload Me
End Sub

Sub DrawRoughMaze()
ScaleMode = 3
PB.Move 0, 0, Maze.Size, Maze.Size
ScaleMode = 1
PB.ScaleWidth = Maze.Size
PB.ScaleHeight = Maze.Size
VPB.ScaleWidth = Maze.Size
VPB.ScaleHeight = Maze.Size
For i = 0 To Maze.Size
For i2 = 0 To Maze.Size
    If Maze.Object(i, i2) = WALL Then
    PB.PSet (i, i2), vbWhite
    Else
    PB.PSet (i, i2), vbBlack
    End If
Next
Next
PB.Picture = PB.Image
VPB.Cls
VPB.PaintPicture PB.Picture, 0, 0, VPB.Width, VPB.Height, 0, 0, PB.Width, PB.Height
For i = 0 To Maze.Size
VPB.Line (0, i)-(Maze.Size, i), vbGreen
VPB.Line (i, 0)-(i, Maze.Size), vbGreen
Next
End Sub

Sub MoveCursor(x As Byte, y As Byte)
Cursor.x = x
Cursor.y = y
Box.Move x, y, 1, 1
End Sub

