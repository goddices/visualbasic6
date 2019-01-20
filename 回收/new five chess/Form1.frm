VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   5715
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   600
      ScaleHeight     =   4470
      ScaleWidth      =   4470
      TabIndex        =   0
      Top             =   360
      Width           =   4500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const DegreeWidth = 9
Private x_basic As Single, y_basic As Single

Private Sub Form_Activate()
   x_basic = Picture1.Width / DegreeWidth
    y_basic = Picture1.Height / DegreeWidth
    Picture1.DrawWidth = 1.2
    For i = 0 To DegreeWidth
        m = i * x_basic + 0.5 * x_basic
             
        Picture1.Line (m, 0.5 * y_basic)-(m, Picture1.Height - 0.5 * y_basic)
        Picture1.Line (0.5 * x_basic, m)-(Picture1.Width - 0.5 * x_basic, m)
    Next i
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then

    a = Int(X / x_basic)
    b = Int(Y / y_basic)
    
    MsgBox a & "  " & b
End If
End Sub

