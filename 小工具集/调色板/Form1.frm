VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "调色板"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7260
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7260
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   3000
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   2
      Left            =   3960
      Max             =   255
      TabIndex        =   5
      Top             =   1800
      Width           =   3015
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   1
      Left            =   3960
      Max             =   255
      TabIndex        =   4
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   0
      Left            =   3960
      Max             =   255
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   360
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "颜色值(for vb)"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "网页颜色值"
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "红                 绿                蓝"
      Height          =   1815
      Left            =   2760
      TabIndex        =   8
      Top             =   465
      Width           =   375
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      Height          =   375
      Index           =   0
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub HScroll1_Change(Index As Integer)
    Picture1_Click
    Label1_Click (Index)
End Sub

 
Private Sub HScroll1_Scroll(Index As Integer)
    HScroll1_Change (Index)
End Sub


Private Sub Label1_Click(Index As Integer)
    Label1(Index).Caption = Hex(CStr(HScroll1(Index).Value))
End Sub

Private Sub Picture1_Click()
    Dim R As Long, G As Long, B As Long
    R = HScroll1(0).Value
    G = HScroll1(1).Value
    B = HScroll1(2).Value
    Picture1.BackColor = RGB(R, G, B)
    Text1.Text = Hex(CStr(R * 256 * 256 + G * 256 + B))
    Text2.Text = Hex(CStr(RGB(R, G, B)))
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo err
    Dim c(2) As Long
    Dim color As Long
    If KeyCode = 13 Then '
        color = CLng("&h" + Text1.Text)
        c(0) = color \ &H10000          'R
        c(1) = color \ &H100 Mod &H100  'G
        c(2) = color Mod &H100          'B
        For i = 0 To 2
            HScroll1(i).Value = c(i)
        Next
    End If
    Exit Sub
err:
    MsgBox "输入的数据类型不正确，确保输入的数值是大于等于 0 或小于等于 0FFFFH 的十六进制数(十进制16777215D)", vbExclamation, "错误"
End Sub
