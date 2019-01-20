VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   6600
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Logistic()
    Dim Xn As Double
    
    Me.Line (0, -0.01)-(5, -0.01)
    Me.Line (-0.01, 0)-(-0.01, 1)
    
    For k = 0 To 4 Step 0.01
    Xn = 0.5
        For i = 1 To 300
        
            Xn = Xn * k * (1 - Xn)
            
            PSet (k, Xn), vbRed
        Next
        
    Next
End Sub

Private Sub Form_Click()
    Logistic
End Sub

Private Sub Form_Load()
    Me.Scale (-0.5, 1.2)-(4.5, -0.1)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label1.Caption = "x = " & x & vbNewLine & "y = " & y
End Sub

 
 
