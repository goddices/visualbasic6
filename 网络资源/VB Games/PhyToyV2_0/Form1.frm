VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Physics Toy V2.0        www.vbgamedev.com"
   ClientHeight    =   10305
   ClientLeft      =   1425
   ClientTop       =   2250
   ClientWidth     =   13425
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   7.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   687
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   895
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606060&
      Height          =   9900
      Left            =   120
      ScaleHeight     =   -18.346
      ScaleLeft       =   -2
      ScaleMode       =   0  'User
      ScaleTop        =   15
      ScaleWidth      =   24.457
      TabIndex        =   0
      Top             =   120
      Width           =   13200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Mouse_pos.X = X
    Mouse_pos.Y = Y
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    force_CE = True
    Mouse_pos.X = X
    Mouse_pos.Y = Y

    MouseHit_Num = -1
    Dim i As Long, vno As D3DVECTOR

    '判断用户选择了哪个刚体
    For i = 1 To NUMBox

        If box(i).TAPE = TAPEBOX Then
            If pntINBOX(Mouse_pos, box(i), vno) <> 0 Then
                MouseHit_Num = i
                Exit For
            End If

        ElseIf box(i).TAPE = TAPECIRCLE Then
            If VDst(box(i).pos, Mouse_pos) < box(i).Rbou Then
                MouseHit_Num = i
                Exit For
            End If

        End If

    Next

    If MouseHit_Num <> -1 Then
        MouseHit_pos = ProjectionWorldPnt_Body(Mouse_pos, box(MouseHit_Num))
    End If

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    force_CE = False
End Sub

Public Sub DrawLine(V1 As D3DVECTOR, V2 As D3DVECTOR, Optional LCOLOR As Long = 0)
    Picture1.Line (V1.X, V1.Y)-(V2.X, V2.Y), LCOLOR
End Sub

Public Sub DrawCircle(V1 As D3DVECTOR, R As Single, Optional LCOLOR As Long = 0)
    Picture1.Circle (V1.X, V1.Y), R, LCOLOR
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
