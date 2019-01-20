VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   323
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1320
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   3
      Top             =   3480
      Width           =   300
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   2280
      Top             =   2880
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3000
      Top             =   2760
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   3480
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   960
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   1
      Top             =   2760
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   720
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   270
      ScaleWidth      =   1890
      TabIndex        =   0
      Top             =   960
      Width           =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Dim i As Integer
Dim d As Integer

Private Sub Form_Load()
i = 0
d = 2
End Sub

Private Sub Timer1_Timer()
i = i + 1
 If i > 7 Then i = 1
 
 
BitBlt Picture2.hDC, 0, 0, i * 18, i * 18, Picture1.hDC, (i - 1) * 18, 0, vbSrcCopy

End Sub

Private Sub Timer2_Timer()
d = d + 1
 If d > 7 Then d = 1
 
 
BitBlt Picture3.hDC, 0, 0, d * 18, d * 18, Picture1.hDC, (d - 1) * 18, 0, vbSrcCopy

End Sub
