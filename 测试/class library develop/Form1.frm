VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   1440
      TabIndex        =   3
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "->"
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1320
      TabIndex        =   1
      Top             =   4440
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   480
      ScaleHeight     =   1515
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<-"
      Height          =   735
      Left            =   2760
      TabIndex        =   4
      Top             =   4440
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mcls As New MyMap
Private b1 As Block
Private b2 As Block
 
Private Sub Command1_Click()
mcls.MainPictureBox = Picture1

Debug.Print Picture1.Height
Debug.Print Picture1.Width

End Sub

 
 

Private Sub Command2_Click()
 b2.PointX = b2.PointX + 1
b1.PointX = b1.PointX + 1

End Sub

 

 
Private Sub Command3_Click()
Set b1 = New Block
 Set b2 = New Block
 b1.PointX = 7
 
 b2.PointX = 2 '

End Sub

Private Sub Command4_Click()

b1.PointX = b1.PointX - 1
 b2.PointX = b2.PointX - 1
End Sub

Private Sub Command5_Click()
Dim a As Integer
a = b1.PointX + 1
End Sub

Private Sub Command6_Click()
b1.PointX = b1.PointX + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mcls = Nothing
Set b1 = Nothing
Set b2 = Nothing
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgBox intCoordinates()
End Sub
