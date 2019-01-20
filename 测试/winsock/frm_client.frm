VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
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
   Begin MSWinsockLib.Winsock sck_c_UDP 
      Left            =   4200
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "join"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2040
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox txt_ip_s 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Text            =   "192.168.198.1"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "connect"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock sck_c 
      Left            =   3480
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1002
      LocalPort       =   1001
   End
   Begin VB.Label Label1 
      Caption         =   "IP address of the server"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
sck_c.RemoteHost = txt_ip_s.Text
sck_c.Connect
sck_c_UDP.SendData sck_c.LocalIP
End Sub

Private Sub sck_c_Connect()
MsgBox "Connection Successed", vbOKOnly + vbInformation
End Sub

 
