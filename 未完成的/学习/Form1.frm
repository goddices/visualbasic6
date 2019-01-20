VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   7230
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1463
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   743
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登录"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "密码输入要注意大小写"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Left            =   4080
      TabIndex        =   5
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   4080
      TabIndex        =   4
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "注册"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   6120
      TabIndex        =   3
      Top             =   2280
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
Private Const IDC_HAND = "#32649"
Private hHandCursor     As Long


Private Sub Form_Load()
hHandCursor = LoadCursor(0, IDC_HAND)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbBlue
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor hHandCursor
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor hHandCursor
Label1.ForeColor = vbRed
End Sub

