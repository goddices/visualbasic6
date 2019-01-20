VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "字符串 转换到 ASCII"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4245
   StartUpPosition =   1  '所有者中心
   Begin VB.OptionButton OptEN 
      Caption         =   "ASCII"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   2200
      Width           =   855
   End
   Begin VB.OptionButton OptCN 
      Caption         =   "Unicode"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   2200
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   2
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "转换"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
On Error GoTo err1
    Dim strText As String
    Dim temp As String
    Dim strlen As Long
    Dim rStr As String
    Dim a As Integer
    Dim i As Integer
    
    Text2.Text = ""
    strText = Text1.Text
    strlen = Len(strText)
     
    If OptCN.Value = True Then
        For i = 1 To strlen
            temp = Mid(strText, i, 1)
            rStr = rStr & " " & Hex(AscW(temp))
        Next
    ElseIf OptEN.Value = True Then
        For i = 1 To strlen
            a = Asc(Mid(strText, i, 1))
            rStr = rStr & " " & Hex(a)
        Next
    Else
        rStr = "请选择英文或中文"
    
    End If
    Text2.Text = rStr
    Exit Sub
err1:
    Text2.Text = "模式选择错误"
End Sub

 

Private Sub Form_Load()
    Text1.ForeColor = &HDDDFDB
End Sub

Private Sub Text1_GotFocus()
    Text1.Text = ""
    Text1.ForeColor = 0
End Sub
