VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logistic Mapping"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   570
   StartUpPosition =   3  '窗口缺省
   Begin VB.VScrollBar VS1 
      Height          =   2775
      Left            =   7680
      Max             =   400
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   " X[n+1]=X[n]·k·( 1 - X[n] )"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "迭代次数"
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Xn"
      Height          =   180
      Left            =   840
      TabIndex        =   3
      Top             =   1440
      Width           =   180
   End
   Begin VB.Label Label2 
      Caption         =   "k"
      Height          =   255
      Left            =   7560
      TabIndex        =   2
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub Logistic(k As Single)
    Dim Xn As Double
    Xn = 0.5
    Me.Line (0, -0.01)-(300, -0.01)
    Me.Line (-0.01, 0)-(-0.01, 1)
    For i = 1 To 300
        Xn = Xn * k * (1 - Xn)
        PSet (i, Xn), vbRed
    Next
End Sub

Private Sub Form_Load()
    Me.Scale (-60, 2)-(430, -0.5)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Label1.Caption = "x = " & x & vbNewLine & "y = " & y
End Sub

Private Sub VS1_Change()
    Dim val As Single
     
    val = VS1.Value * 0.01
    Label2.Caption = "k = " & CStr(val)
    Form1.Cls
    Logistic val
End Sub

Private Sub VS1_Scroll()
    VS1_Change
End Sub
