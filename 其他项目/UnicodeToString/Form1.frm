VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   5985
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   2535
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2760
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim org As String
Dim re() As String
Dim txtout As String
Dim lcode As Long
Private Sub Command1_Click()
    org = Text1.Text
    org = Replace(org, vbCrLf, "")
    org = Trim(org)
    
    re = Split(org, " ")
    
    Text1.Text = org
    
   For i = 0 To UBound(re)
      re(i) = Left("&H" & re(i), 6)
      txtout = txtout & ChrW(re(i))
   Next
   Text2.Text = txtout
    
End Sub

