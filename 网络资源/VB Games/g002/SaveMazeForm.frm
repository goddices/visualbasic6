VERSION 5.00
Begin VB.Form SaveLevelForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Level"
   ClientHeight    =   4068
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   3024
   Icon            =   "SaveMazeForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4068
   ScaleWidth      =   3024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox FileNameT 
      Height          =   288
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   2772
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   1332
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   492
      Left            =   1560
      TabIndex        =   1
      Top             =   3480
      Width           =   1332
   End
   Begin VB.FileListBox File1 
      Height          =   2760
      Left            =   120
      Pattern         =   "*.sml"
      TabIndex        =   0
      Top             =   120
      Width           =   2772
   End
   Begin VB.Label Label1 
      Caption         =   "Filename :"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   2772
   End
End
Attribute VB_Name = "SaveLevelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Select Case cmdSave.Caption
  Case "Load": LoadIt
  Case "Save": SaveIt
  Case "Delete": DeleteIt
End Select
End Sub

Private Sub DeleteIt()
If MsgBox("Are you sure you want to delete " & FileNameT & "?", vbOKCancel + vbQuestion, "Delete Level?") Then DeleteFile File1.Path & "\" & File1.FileName
End Sub

Private Sub LoadIt()
If File1.FileName = "" Then
  MsgBox "Please select a file first", vbOKOnly + vbExclamation, "Cannot save without a file name"
  Exit Sub
End If

If LoadMaze(File1.Path & "\" & File1.FileName) Then
  MsgBox "Loaded " & FileNameT & " OK", , FileNameT & " Loaded OK"
Else
  MsgBox "There was an error when attempting to load " & FileNameT, vbOKOnly + vbExclamation, "Error - level not loaded!!!"
End If

DForm.WallPic = LoadPicture(App.Path & "\Resources\Pictures\Walls\" & Maze.WallType & ".ico")
DForm.FloorPic = LoadPicture(App.Path & "\Resources\Pictures\Floors\" & Maze.FloorType & ".ico")

Unload Me
End Sub

Private Sub SaveIt()
If FileNameT.Text = "" Then
  MsgBox "Please type in a file name first", vbOKOnly + vbExclamation, "Cannot save without a file name"
  Exit Sub
End If

If SaveMaze(File1.Path & "\" & FileNameT & ".sml") Then
  MsgBox "Saved " & FileNameT & " OK", , FileNameT & " OK"
Else
  MsgBox "There was an error when attempting to save " & FileNameT, vbOKOnly + vbExclamation, "Error - level not saved!!!"
End If

Unload Me
End Sub

Private Sub File1_Click()
FileNameT = (Left(File1.FileName, Len(File1.FileName) - 4))
Label1 = "File name : " & FileNameT
End Sub

Private Sub FileNameT_Change()
Label1 = "File name : " & FileNameT
End Sub

Private Sub Form_Load()
File1.Path = App.Path & "\Resources\Levels\"
End Sub

Private Sub Form_Unload(Cancel As Integer)
DForm.DrawRoughMaze
End Sub
