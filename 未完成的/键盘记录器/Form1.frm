VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Left            =   1440
      Top             =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, &H80
     'Me.Caption = "¼üÅÌ¼àÊÓÆ÷"
     SetTimer Me.hwnd, 0, 1, AddressOf TimerProc
     Timer1.Interval = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
     KillTimer Me.hwnd, 0
End Sub

Private Sub Timer1_Timer()
Text1.Text = sSave
End Sub



