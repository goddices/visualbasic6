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
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Const WM_CLOSE = &H10


Private Sub Command1_Click()
'Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE" & Space(1) & "www.suwumuyangzjd.com/message.asp", vbMinimizedNoFocus
Dim hwnd1 As Long, str As String
str = String(255, Chr(0))
hwnd1 = FindWindow(vbNullString, "QQ2009 正式版")
GetClassName hwnd1, str, 255

Print "Handle to the window of QQ.exe : " & hwnd1
Print "Class name of the window : " & str
SendMessage hwnd1, WM_CLOSE, 0, 0
End Sub
