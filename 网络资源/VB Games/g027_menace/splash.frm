VERSION 5.00
Begin VB.Form splash 
   Caption         =   "Menace"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Menace Map Editor"
      Height          =   375
      Left            =   135
      TabIndex        =   1
      Top             =   2625
      Width           =   1830
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play Menace"
      Default         =   -1  'True
      Height          =   420
      Left            =   135
      TabIndex        =   0
      Top             =   2145
      Width           =   1830
   End
   Begin VB.Label Label5 
      Caption         =   "R - restarts the level if you are stuck."
      Height          =   390
      Left            =   2130
      TabIndex        =   6
      Top             =   2595
      Width           =   2340
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   135
      Picture         =   "splash.frx":0000
      Top             =   150
      Width           =   1830
   End
   Begin VB.Label Label4 
      Caption         =   "The Control key jumps menace, Shift and Control will push trollies high up."
      Height          =   690
      Left            =   2130
      TabIndex        =   5
      Top             =   1890
      Width           =   2550
   End
   Begin VB.Label Label3 
      Caption         =   "The aim of the game is to push trollies containing the same shape together into groups of three."
      Height          =   690
      Left            =   2160
      TabIndex        =   4
      Top             =   405
      Width           =   2445
   End
   Begin VB.Label Label2 
      Caption         =   "Quick Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2145
      TabIndex        =   3
      Top             =   150
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "The left and right arrows walk menace.  By holding down shift you can push trollies"
      Height          =   675
      Left            =   2145
      TabIndex        =   2
      Top             =   1155
      Width           =   2460
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
frmBlt.Show
End Sub

Private Sub Command2_Click()
Form1.Show
End Sub

