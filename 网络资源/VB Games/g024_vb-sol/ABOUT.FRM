VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About VB Solitaire"
   ClientHeight    =   2595
   ClientLeft      =   1815
   ClientTop       =   3630
   ClientWidth     =   3615
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2595
   ScaleWidth      =   3615
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      Picture         =   "ABOUT.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label lblVB 
      Alignment       =   2  'Center
      Caption         =   "Visual Basic"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblSol 
      Alignment       =   2  'Center
      Caption         =   "Solitaire"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

  'Unload About Box.
  Unload frmAbout

End Sub

Private Sub Form_Load()

  'Set copyright message in label -- Chr$(169) is copyright symbol.
  lblCopyright.Caption = "Copyright " + Chr$(169) + " 1996 Michael Bertrand"

End Sub

