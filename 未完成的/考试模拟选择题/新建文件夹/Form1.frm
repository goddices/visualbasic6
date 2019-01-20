VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   9420
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check4"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d As String

Private Sub Command1_Click()
d = "ABC"
   For i = 1 To 4
    If d <> "" Then
        Print Left(d, i)
        Print Mid(d, i)
    End If
    
   Next
End Sub

Sub re()

End Sub
