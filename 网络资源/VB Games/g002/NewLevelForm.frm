VERSION 5.00
Begin VB.Form NewLevelForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Level"
   ClientHeight    =   1104
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   3744
   Icon            =   "NewLevelForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1104
   ScaleWidth      =   3744
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   2280
      TabIndex        =   2
      Top             =   600
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1200
      TabIndex        =   0
      Text            =   "20"
      Top             =   240
      Width           =   2292
   End
   Begin VB.Label Label1 
      Caption         =   "Size of level :"
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1812
   End
End
Attribute VB_Name = "NewLevelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ChosenWall As Byte

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo MuffUp:
DForm.CreateBorderedMaze Int(Val(Text1.Text)), ChosenWall
Unload Me
Exit Sub
MuffUp:
MsgBox "Illegal maze size entered, please choose an integer between 10 and 200", vbOKOnly + vbExclamation, "Muff Up!"
End Sub
