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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock sck_s 
      Left            =   3600
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1001
      LocalPort       =   1002
   End
   Begin MSWinsockLib.Winsock sck_c 
      Left            =   4200
      Top             =   2520
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
Private Sub Command1_Click()
MsgBox sck_s.RemoteHostIP
End Sub

Private Sub Form_Unload(Cancel As Integer)
If sck_s.State <> sckClosed Then sck_s.Close
End Sub

Private Sub sck_s_ConnectionRequest(ByVal requestID As Long)
sck_s.Accept requestID
MsgBox "accept request from Id" & requestID
End Sub

Private Sub sck_s_DataArrival(ByVal bytesTotal As Long)
Dim str As String
sck_s.GetData str
MsgBox str
End Sub
