VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "target window"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "清楚"
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "鼠标左键点击左边的方块，窗体将打印 image1 left down,右击打印image2 right down，对于点击右边的方块也是一样"
      Height          =   1095
      Left            =   5760
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Height          =   615
      Left            =   5880
      TabIndex        =   1
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   6960
      Top             =   120
      Width           =   975
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   5760
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Cls
End Sub

Private Sub Form_Activate()
    Print "image1 left    " & Image1.Left
    Print "image1 top      " & Image1.Top
    Print "image2 left    " & Image2.Left
    Print "image2 top      " & Image2.Top
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = X & vbNewLine & Y
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Print "image1 left button down"
    Else
        Print "image1 right button down"
    End If
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Print "image2 left button down"
    Else
        Print "image2 right button down"
    End If
End Sub
