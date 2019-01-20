VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   3660
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private arr()

Private Sub Command1_Click()
    Dim hwnd1 As Long
    Dim t As String
    
    t = Space(255)
    
      For i = 0 To 3
        hwnd1 = FindWindow(vbNullString, arr(i))
        GetClassName hwnd, t, 255
        List1.AddItem t
        SetWindowText hwnd1, "SB"
     Next
    
End Sub

Private Sub Form_Load()
arr = Array("我的电脑", "我的文档", "计算机", "web")
End Sub
