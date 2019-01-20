VERSION 5.00
Begin VB.Form AEditBookId 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "正在修改借书人员"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "AEditBookId.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtBookId 
         Height          =   270
         Left            =   1080
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtBookIdName 
         Height          =   270
         Left            =   3480
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtBookIdClass 
         Height          =   270
         Left            =   1080
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtBookIdDepart 
         Height          =   270
         Left            =   3480
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtBookIdZhi 
         Height          =   270
         Left            =   1080
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Left            =   1080
         TabIndex        =   3
         Top             =   2160
         Width           =   3855
      End
      Begin VB.CommandButton cmdOkCancel 
         Caption         =   "保存(&E)"
         Default         =   -1  'True
         Height          =   375
         Index           =   0
         Left            =   3960
         TabIndex        =   2
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdOkCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   1
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "借书证号"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "姓    名"
         Height          =   180
         Index           =   1
         Left            =   2640
         TabIndex        =   12
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "班    级"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "部    门"
         Height          =   180
         Index           =   3
         Left            =   2640
         TabIndex        =   10
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "职    称"
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "AEditBookId.frx":0442
         Top             =   2280
         Width           =   480
      End
   End
End
Attribute VB_Name = "AEditBookId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOkCancel_Click(Index As Integer)
Select Case Index
    Case 0
        mAddEditId = txtBookId & vbNullString
        mAddEditIdName = txtBookIdName & vbNullString
        mAddEditIdClass = txtBookIdClass & vbNullString & " "
        mAddEditIdDepart = txtBookIdDepart & vbNullString & " "
        mAddEditIdZhi = txtBookIdZhi & vbNullString & " "
        mSave = True
        Unload Me
    Case 1
        mSave = False
        Unload Me
End Select
End Sub
Private Sub Form_Load()
txtBookId.Text = mAddEditId
txtBookIdName = mAddEditIdName
txtBookIdClass = mAddEditIdClass
txtBookIdDepart = mAddEditIdDepart
txtBookIdZhi = mAddEditIdZhi
End Sub

