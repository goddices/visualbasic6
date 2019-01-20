VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Dim str As String
    
    Dim r As Variant
    
    Open App.Path & "\target.exe" For Binary As #1
    Open App.Path & "\added.exe" For Binary As #2
    ReDim ar(LOF(2)) As Byte
     
     
     
    'Get #1, LOF(1), ar
     
    For i = 1 To LOF(2)
        Get #2, , ar(i)
        Put #1, LOF(1) + 1, ar(i)
    Next
    
    Close #1, #2
End Sub
