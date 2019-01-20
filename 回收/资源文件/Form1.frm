VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   5625
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Height          =   735
      Left            =   600
      ScaleHeight     =   675
      ScaleWidth      =   3675
      TabIndex        =   2
      Top             =   360
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   3720
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   720
      ScaleHeight     =   1875
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub Command1_Click()
 
  
 Picture1.Picture = LoadResPicture(101, 0)
BitBlt Picture2.hDC, 0, 0, 126, 126, Picture1.hDC, 0, 0, vbSrcCopy
 End Sub
