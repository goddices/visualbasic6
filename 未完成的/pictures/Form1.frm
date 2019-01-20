VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   600
      Left            =   3240
      TabIndex        =   0
      Top             =   2640
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim minX As Integer, minY  As Integer
Dim cX  As Integer, cY As Integer
Private Sub Command1_Click()
Dim i As Integer
Dim j As Integer
Dim index As Integer
Dim pic(9) As Picture
For i = 0 To 2
    For j = 0 To 2
        index = index + 1
        Set pic(index) = LoadPicture(App.Path & "\" & index & ".bmp")
        Me.PaintPicture pic(index), i * 42 + 2, j * 42 + 2, 40, 40
        
    Next
Next
End Sub


Private Sub Form_Load()
 
minX = Form1.Width / 95
minY = Form1.Height / 95
 MsgBox minX
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
cX = Int(X / minX)
cY = Int(Y / minY)
 
cX = cX * minX + 0.5 * minX
cY = cY * minY + 0.5 * minY
Form1.Line (cX - 20, cY - 20)-(cX + 21, cY + 21), vbRed, B
End Sub

Sub border()
Form1.DrawWidth = 2


End Sub

 
