VERSION 5.00
Begin VB.Form Aboutfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "Aboutfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0000C000&
      Cancel          =   -1  'True
      Caption         =   "�ر�"
      Height          =   495
      Left            =   4080
      Picture         =   "Aboutfrm.frx":0442
      TabIndex        =   3
      ToolTipText     =   "�ر�"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Label Label3 
      Caption         =   "    δ��������������������ҵ��;����������߱���׷���������ε�Ȩ����"
      Height          =   855
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "·�೬"
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
      Height          =   240
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Ҷͼ�����ϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2040
   End
End
Attribute VB_Name = "Aboutfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
Unload Me
End Sub

