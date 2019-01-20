VERSION 5.00
Begin VB.Form HelpForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "How to play The Haunted Maze"
   ClientHeight    =   1932
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   3936
   Icon            =   "HelpForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1932
   ScaleWidth      =   3936
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   492
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "Controls : Use the arrow keys to move around. Press P to pause and unpause the action"
      Height          =   492
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3732
   End
   Begin VB.Label Label1 
      Caption         =   "The aim: The aim of the game is to collect all the bones in the level without any ghost touching you."
      Height          =   492
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3732
   End
End
Attribute VB_Name = "HelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
