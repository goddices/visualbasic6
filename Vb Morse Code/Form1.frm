VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "摩尔斯电码转换"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7635
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear Both"
      Height          =   855
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox TxtMorse 
      Height          =   2295
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Form1.frx":0000
      Top             =   2880
      Width           =   7095
   End
   Begin VB.CommandButton CpyMorse 
      Caption         =   "Copy Morse"
      Height          =   855
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton CpyEn 
      Caption         =   "Copy English"
      Height          =   855
      Left            =   3240
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton SwitchM2E 
      Caption         =   "Morse To English"
      Height          =   855
      Left            =   4560
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton SwitchE2M 
      Caption         =   "English To Morse"
      Height          =   855
      Left            =   5880
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox TxtEn 
      Height          =   1575
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0012
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SwitchE2M_Click()

End Sub

Private Sub SwitchM2E_Click()

End Sub
