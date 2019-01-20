VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   469
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command3 
      Caption         =   "¡ú"
      Height          =   735
      Left            =   5640
      TabIndex        =   3
      Top             =   3840
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   4815
      Left            =   240
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¡ý"
      Height          =   735
      Left            =   4920
      TabIndex        =   0
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "¡û"
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "¡ü"
      Height          =   735
      Left            =   4920
      TabIndex        =   4
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1095
      Left            =   4320
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VT As Class1


Private Sub Command4_Click()
 VT.PointY = VT.PointY - 1
End Sub

Private Sub Form_Load()
Set VT = New Class1
 
VT.MainDrawingPictureBox = Picture1
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
   Label1.Caption = VT.PointX & vbNewLine & VT.PointY
    Select Case KeyCode
        Case 37
            VT.PointX = VT.PointX - 1
        Case 38
            VT.Transform = VT.Transform + 1
        Case 39
            VT.PointX = VT.PointX + 1
        Case 40
            VT.PointY = VT.PointY + 1
    End Select
    
 
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    a = Int(X / dlLenX)
    b = Int(Y / dlLenY)
    
    MsgBox a & "  " & b & "  " & Coordinates(a, b)
     
End Sub
