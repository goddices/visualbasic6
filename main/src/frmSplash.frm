VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3855
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   5805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3855
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3480
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "请点击这里进入……"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "华文行楷"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "@隶书"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim j As Integer
Dim i As Integer
Label3.Caption = ""
Label2.Caption = "正在初始化系统，请稍后……"
For i = 1 To 3000
ProgressBar2.Value = i
Label1.Caption = Format(ProgressBar2.Value / ProgressBar2.Max, "##%")
DoEvents
Next i
 If Label1.Caption = "100%" Then
   Label2.Caption = "正在更新系统信息库……"
    Label2.ForeColor = &HFFFF00
    For j = 1 To 3000
    ProgressBar2.Value = j
   Label1.Caption = Format(ProgressBar2.Value / ProgressBar2.Max, "##%")
  DoEvents
Next j


LoginSys.Show
Me.Hide
'LoginSys.Option1.Value = True
End If

End Sub

Private Sub Form_Load()
ProgressBar2.Min = 0
ProgressBar2.Max = 3000
Label1.Caption = ""
End Sub

