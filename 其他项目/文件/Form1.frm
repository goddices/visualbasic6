VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   5535
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim FileName As String
  Dim buff()     As Byte
  Dim fl As Long

Private Sub Command1_Click()
    Dim i As Long
    fl = FileLen(FileName)
    ReDim buff(fl - 1)
    Open FileName For Binary As #1
    Get #1, , buff
    Close #1
    
    For i = 0 To fl - 1
        Text1.Text = Text1.Text & Format(Hex(buff(i)), "0") & " "
        If (i + 1) Mod 16 = 0 Then Text1.Text = Text1.Text & vbNewLine
    Next
End Sub

 

Private Sub Command2_Click()
    
End Sub

Private Sub Form_Load()
  FileName = App.Path & "\ewh.db"
End Sub
 
