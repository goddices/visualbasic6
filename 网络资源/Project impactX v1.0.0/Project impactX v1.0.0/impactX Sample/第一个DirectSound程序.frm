VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��һ��DirectSound����"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4200
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdStop 
      Caption         =   "ֹͣ"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "����"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'impactX Game Engine
'written by Davy.xu
'һ������WAV��DX����
'���ڲ˵� ����->��������� DirectX 8 for Visual Basic Type Library
Dim xa As New xAudio
Dim Mywav As DirectSoundSecondaryBuffer8

Private Sub cmdPlay_Click()
    xa.PlayWav Mywav
End Sub

Private Sub cmdStop_Click()
    xa.StopWav Mywav
End Sub

Private Sub Form_Load()
    xa.InitDXSound Me.hWnd
    Me.Show
    Set Mywav = xa.LoadWav("explode.wav") ' ����WAV
End Sub

Private Sub Form_Unload(Cancel As Integer)
    xa.UnloadDXSound
    xa.ReleaseWav Mywav
End Sub
