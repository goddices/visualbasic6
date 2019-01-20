VERSION 5.00
Begin VB.Form AboutForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Haunted Maze 2"
   ClientHeight    =   2616
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   3744
   Icon            =   "AboutForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2616
   ScaleWidth      =   3744
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   492
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   1212
   End
   Begin VB.Image Image1 
      Height          =   732
      Left            =   120
      Picture         =   "AboutForm.frx":030A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   732
   End
   Begin VB.Label Label3 
      Caption         =   "By Simon Price"
      Height          =   372
      Left            =   1200
      TabIndex        =   2
      Top             =   1560
      Width           =   2052
   End
   Begin VB.Label Label2 
      Caption         =   "Version 2.0 "
      Height          =   252
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   1932
   End
   Begin VB.Label Label1 
      Caption         =   "The Haunted Maze "
      Height          =   252
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   1932
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
