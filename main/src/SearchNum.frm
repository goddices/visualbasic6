VERSION 5.00
Begin VB.Form SearchNum 
   Caption         =   "查找图书编号"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   5220
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOkCancel 
      Cancel          =   -1  'True
      Caption         =   "不想查了(&E)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdOkCancel 
      Caption         =   "填好了(&E)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtBookNum 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "SearchNum.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请输入查找图书的编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   2565
   End
End
Attribute VB_Name = "SearchNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOkCancel_Click(Index As Integer)
Select Case Index
    Case 0
        If txtBookNum.Text = "" Then
            MsgBox "你还没填图书编号呢？", 0 + 48, "提示"
            txtBookNum.SetFocus
            Exit Sub
        End If
        BookBianHao = txtBookNum
        SearchFlag = True
        Unload Me
    Case 1
        SearchFlag = False
        Unload Me
End Select
End Sub
Private Sub Form_Load()
txtBookNum.Text = ""
End Sub
