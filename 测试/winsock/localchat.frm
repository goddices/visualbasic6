VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdSend 
      Caption         =   "发送"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Text            =   "1"
      Top             =   3480
      Width           =   4335
   End
   Begin VB.TextBox txtRecord 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.TextBox txtIP 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdconnect 
      Caption         =   "连接"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock Sck_s 
      Left            =   3720
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Sck_c 
      Left            =   3120
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdconnect_Click()
Sck_c.RemoteHost = txtIP.text
Sck_c.Connect
End Sub

Private Sub cmdSend_Click()
Sck_c.SendData Text1.text
txtRecord.text = txtRecord.text & vbNewLine & Text1.text
End Sub

Private Sub Form_Load()
Sck_s.Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
Sck_s.Close
Sck_c.Close
End Sub

Private Sub Sck_c_Connect()
MsgBox "connected"
End Sub

Private Sub Sck_s_ConnectionRequest(ByVal requestID As Long)
Sck_s.Accept requestID
End Sub

Private Sub Sck_s_DataArrival(ByVal bytesTotal As Long)
Dim str As String
Sck_s.GetData str
txtRecord.text = txtRecord.text & vbNewLine & str

End Sub

