VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirectShow����"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4455
   StartUpPosition =   3  '����ȱʡ
   Begin VB.HScrollBar vol 
      Height          =   255
      Left            =   1080
      Max             =   100
      TabIndex        =   7
      Top             =   4080
      Value           =   100
      Width           =   1935
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "��ͣ"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "ֹͣ"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "����"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   3960
   End
   Begin VB.HScrollBar pgs 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   2895
   End
   Begin VB.HScrollBar bal 
      Height          =   255
      Left            =   1080
      Max             =   100
      Min             =   -100
      TabIndex        =   1
      Top             =   3720
      Width           =   1935
   End
   Begin VB.PictureBox pic 
      Height          =   2850
      Left            =   120
      ScaleHeight     =   1715.534
      ScaleMode       =   0  'User
      ScaleWidth      =   2919.899
      TabIndex        =   0
      Top             =   120
      Width           =   4200
   End
   Begin VB.Label Label3 
      Caption         =   "����:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "����ƽ��:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "����λ��"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Impact Game Engine
'written by Davy.xu
'һ���򵥵Ķ�ý���ļ�������
'���ڲ˵� ����->��������� DirectX 8 for Visual Basic Type Library��ActiveMovie control
Dim xs As New xShow
Private Sub Bal_Change()
    xs.Balance = bal.Value
    Me.Caption = bal.Value
End Sub

Private Sub cmdPause_Click()
    xs.PauseMedia
End Sub

Private Sub cmdPlay_Click()
    xs.PlayMedia
End Sub

Private Sub cmdStop_Click()
    xs.StopMedia
End Sub

Private Sub Form_Load()
    xs.InitDXShow pic.hWnd, pic.Width / 15, pic.Height / 15 '(���ڿ����������Ļ���������15��)
    Me.Show
    DoEvents
    xs.LoadMedia "CLOCKTXT.avi" '��������MP3���Լ��������ԣ�"
    xs.PlayMedia
    pgs.Max = xs.Duration
End Sub


Private Sub pgs_Scroll()
    xs.MediaPosition = pgs.Value
End Sub

Private Sub Timer1_Timer()
    pgs.Value = xs.MediaPosition
    Label1.Caption = pgs.Value
End Sub

Private Sub vol_Change()
    xs.Volume = vol.Value
End Sub
