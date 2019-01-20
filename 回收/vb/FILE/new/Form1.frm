VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   0
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommDlg1 
      Left            =   0
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim FileStr As String
Dim str As String
Dim h As String
Dim ar As Byte
CommDlg1.Filter = "所图片文件|*.jpg;*.bmp;*.gif"
CommDlg1.ShowOpen

FileStr = CommDlg1.FileName
If FileStr <> "" Then
    Open FileStr For Binary As #1
    For i = 1 To 100
        Get #1, i, ar
        h = "0x" & Format(Hex(ar), "00")
        str = str & h & vbNewLine
    Next
    
    str = str & "......." & vbNewLine
    
    For i = LOF(1) - 100 To LOF(1)
        Get #1, i, ar
        h = "0x" & Format(Hex(ar), "00")
        str = str & h & vbNewLine
    Next
    
    Close #1
End If

Text1.Text = str

End Sub
