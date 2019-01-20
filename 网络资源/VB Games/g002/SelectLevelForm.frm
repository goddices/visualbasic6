VERSION 5.00
Begin VB.Form SelectLevelForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select a Level"
   ClientHeight    =   3528
   ClientLeft      =   2292
   ClientTop       =   1476
   ClientWidth     =   3684
   Icon            =   "SelectLevelForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3528
   ScaleWidth      =   3684
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   492
      Left            =   2280
      TabIndex        =   2
      Top             =   2880
      Width           =   1332
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   492
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1332
   End
   Begin VB.FileListBox File1 
      Height          =   3336
      Left            =   120
      Pattern         =   "*.sml"
      TabIndex        =   0
      Top             =   120
      Width           =   2052
   End
End
Attribute VB_Name = "SelectLevelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdLoad_Click()
If File1.FileName = "" Then
  MsgBox "Please select a level first!", vbOKOnly + vbExclamation, "No level selected!!!"
  Exit Sub
End If

If LoadMaze(File1.Path & "\" & File1.FileName) Then
  GForm.PlaceAllObjects
  GForm.PaintFloor
  GForm.PausedL.Visible = False
  GForm.MoveT.Enabled = True
sndPlaySound App.Path & "\Resources\Sounds\levelstart.wav", &H2
  Unload Me
Else
  MsgBox "Error occured when attempting to load the level!", vbOKOnly + vbExclamation, "Muff Up!!! Level not loaded!"
End If
End Sub

Private Sub Form_Load()
File1.Path = App.Path & "\Resources\Levels"
File1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
File1.Refresh
End Sub
