Attribute VB_Name = "HM2Mod"
Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As Any, ByVal uFlags As Long) As Long

Global Const SND_ASYNC = &H1     ' Play asynchronously
Global Const SND_NODEFAULT = &H2 ' Don't use default sound
Global Const SND_MEMORY = &H4    ' lpszSoundName points to a memory file

Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Type TMaze
  Object() As Byte
  WallType As Byte
  Size As Byte
  NMEcount As Byte
  FloorType As Byte
  LaserCount As Byte
  BoneCount As Byte
  Frame As Byte
End Type

'object type constants
Public Const NOWT = 0
Public Const WALL = 1
Public Const COOLMAN = 2
Public Const NME0 = 3
Public Const NME1 = 4
Public Const NME2 = 5
Public Const NME3 = 6
Public Const NMEangry0 = 7
Public Const NMEangry1 = 8
Public Const NMEangry2 = 9
Public Const NMEangry3 = 10
Public Const NMEconfused0 = 11
Public Const NMEconfused1 = 12
Public Const NMEconfused2 = 13
Public Const NMEconfused3 = 14
Public Const NMEalert0 = 15
Public Const NMEalert1 = 16
Public Const NMEalert2 = 17
Public Const NMEalert3 = 18
Public Const BONE = 19
'total no. of walls
Public Const NUMWALLS = 28
'total no. of floors
Public Const NUMFLOORS = 30

Type TCharacter
  x As Byte
  y As Byte
  dir As Byte
  State As Byte
  Frame As Byte
End Type

'dir (direction) constants
Public Const dUP = 0
Public Const dRIGHT = 1
Public Const dDOWN = 2
Public Const dLEFT = 3
Public Const dFOWARD = 4

'state constants
Public Const NORMAL = 0
Public Const ALERT = 1
Public Const CONFUSED = 2
Public Const ANGRY = 3

'move type constants
Public Const RANDOM = 0

Type TCoOrd
  x As Integer
  y As Integer
End Type

Public Maze As TMaze
Public Man As TCharacter
Public Ghost() As TCharacter

Public i As Integer
Public i2 As Integer
Public i3 As Integer
Public i4 As Integer
Public i5 As Integer
Public TempByte As Byte

Public bMove As Boolean
'temp move direction constants
Public Const HORIZONTAL = True
Public Const VERTICAL = False

Public Temp As TCoOrd
Public Position(1 To 14, 1 To 15) As TCoOrd
Public Lives As Byte

Public Sub CalcPositions()
'this calculates the co-ords of each space in the maze
For i = 1 To 14
For i2 = 1 To 15
 Position(i, i2).x = (i * 24) - (i2 * 8) - 10
 Position(i, i2).y = (i2 * 8) - 16
Next
Next
End Sub

Public Sub CreateMaze(Size As Byte, WallType As Byte)
UnloadMaze 'first unload the old maze
Maze.WallType = WallType
Maze.Size = Size 'put in new size + style
'redeclare array of objects
ReDim Maze.Object(Maze.Size, Maze.Size)
End Sub

Public Sub UnloadMaze() 'unloads all contents of the maze
Maze.Size = 0
Maze.WallType = 0
ReDim Maze.Object(0 To 0, 0 To 0)
End Sub

Public Function SaveMaze(FileName As String) As Boolean
'this saves the maze to the file given
On Error GoTo MuffUp
Open FileName For Random As #1 Len = 1
  Put #1, 1, Maze.Size
  Put #1, 2, Maze.WallType
  Put #1, 3, Maze.FloorType
For i = 0 To Maze.Size
For i2 = 0 To Maze.Size
  Put #1, (Maze.Size * i) + i2 + 4, Maze.Object(i, i2)
Next
Next
Close #1
SaveMaze = True
Exit Function
MuffUp:
SaveMaze = False
End Function

Public Function LoadMaze(FileName As String) As Boolean
'this loads a maze from the file given
On Error GoTo MuffUp:
Open FileName For Random As #1 Len = 1
  Get #1, 1, Maze.Size
  ReDim Maze.Object(Maze.Size, Maze.Size)
  Get #1, 2, Maze.WallType
  Get #1, 3, Maze.FloorType
For i = 0 To Maze.Size
For i2 = 0 To Maze.Size
  Get #1, (Maze.Size * i) + i2 + 4, Maze.Object(i, i2)
Next
Next
Close #1
LoadMaze = True
Exit Function
MuffUp:
LoadMaze = False
End Function

