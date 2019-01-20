VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form setfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "Setfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton cmdOkCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   1
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdOkCancel 
         Caption         =   "确定(&E)"
         Default         =   -1  'True
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   9
         Top             =   2040
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Height          =   135
         Left            =   -74400
         TabIndex        =   8
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtCost 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72600
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   3135
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   3000
         TabIndex        =   4
         Top             =   1320
         Width           =   225
         _ExtentX        =   476
         _ExtentY        =   582
         _Version        =   393216
         BuddyControl    =   "txtLentNum"
         BuddyDispid     =   196614
         OrigLeft        =   2640
         OrigTop         =   1080
         OrigRight       =   2865
         OrigBottom      =   1455
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtLentNum 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1320
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "超出一天罚款金额"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   -74400
         TabIndex        =   6
         Top             =   1320
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "每人借书册数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   600
         TabIndex        =   2
         Top             =   1320
         Width           =   1260
      End
   End
End
Attribute VB_Name = "setfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type MSet
    BookNum As Integer
    BookCost As Single
End Type
Dim SetFlag As MSet
Private Sub cmdOkCancel_Click(Index As Integer)
Select Case Index
    Case 0
        BookNum = Val(txtLentNum)
        FaCost = Val(txtCost)
        SetFlag.BookNum = Val(txtLentNum.Text)
        SetFlag.BookCost = Val(txtCost.Text)
        Put #1, 1, SetFlag
        Unload Me
    Case 1
        Unload Me
End Select
End Sub
Private Sub Form_Load()
Open "Database\Set.Dat" For Random As #2 Len = Len(SetFlag)
Get #1, 1, SetFlag
txtLentNum = SetFlag.BookNum
txtCost = SetFlag.BookCost
End Sub

Private Sub Form_Resize()
MsgBox "请确认图书归还以后在修改这些设置", 0 + 48, "警告"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #2
End Sub

