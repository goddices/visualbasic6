VERSION 5.00
Begin VB.Form LentLogin 
   Caption         =   "��¼"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "LentLogin.frx":0000
   ScaleHeight     =   3555
   ScaleWidth      =   5565
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOkCancel 
      BackColor       =   &H0000C000&
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1560
      Picture         =   "LentLogin.frx":13780
      TabIndex        =   3
      ToolTipText     =   "ȡ������"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOkCancel 
      BackColor       =   &H00FFFF00&
      Caption         =   "ȷ��"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   3120
      Picture         =   "LentLogin.frx":13A8A
      TabIndex        =   2
      ToolTipText     =   "��ʼ����"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtBookId 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   3120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   360
      Picture         =   "LentLogin.frx":13ECC
      Top             =   480
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4440
      Picture         =   "LentLogin.frx":1430E
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "���������֤����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "LentLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rst As Recordset
Private Sub cmdOkCancel_Click(Index As Integer)
Select Case Index
    Case 0
        If txtBookId.Text = "" Then
            MsgBox "���������֤���룡", 0 + 48, "����"
            txtBookId.SetFocus
            Exit Sub
        Else
        BookId = txtBookId.Text
        LoginFlag = True
        Unload Me
        End If
    Case 1
        LoginFlag = False
        Unload Me
End Select
End Sub

Private Sub Form_Load()
txtBookId.Text = ""
Set db = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
'Set rst = db.OpenRecordset("NewBook", dbOpenTable)
End Sub
