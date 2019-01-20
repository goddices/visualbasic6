VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "苏州水费管理系统"
   ClientHeight    =   7920
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9735
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2640
      Top             =   6600
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7545
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "用户名："
            TextSave        =   "用户名："
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu XT 
      Caption         =   "系统"
      Begin VB.Menu DL 
         Caption         =   "登录"
      End
      Begin VB.Menu mmxg 
         Caption         =   "帐户管理"
         Enabled         =   0   'False
      End
      Begin VB.Menu aaa 
         Caption         =   "-"
      End
      Begin VB.Menu TC 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu sfgl 
      Caption         =   "水费管理"
      Enabled         =   0   'False
      Begin VB.Menu jnsf 
         Caption         =   "缴纳水费"
      End
      Begin VB.Menu cxsf 
         Caption         =   "查询缴费情况"
      End
   End
   Begin VB.Menu yfgl 
      Caption         =   "用户管理"
      Enabled         =   0   'False
      Begin VB.Menu bjyf 
         Caption         =   "编辑用户"
      End
      Begin VB.Menu llyf 
         Caption         =   "浏览用户"
      End
   End
   Begin VB.Menu DYFW 
      Caption         =   "打印服务"
      Enabled         =   0   'False
      Begin VB.Menu DRjfqh 
         Caption         =   "当日缴费情况"
      End
   End
   Begin VB.Menu hjsz 
      Caption         =   "环境设置"
      Enabled         =   0   'False
      Begin VB.Menu dqjg 
         Caption         =   "当前水费价格"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bjyf_Click()
Form31.Show
End Sub
'Download by http://down.liehuo.net
Private Sub cxsf_Click()
Form22.Show
End Sub

Private Sub DL_Click()
MDIForm1.mmxg.Enabled = False
MDIForm1.sfgl.Enabled = False
MDIForm1.yfgl.Enabled = False
MDIForm1.DYFW.Enabled = False
MDIForm1.hjsz.Enabled = False
Form11.Show
End Sub

Private Sub dqjg_Click()
Form51.Show
End Sub

Private Sub DRjfqh_Click()
DataReport1.Show
End Sub

Private Sub HELP_Click()
Form13.Show
End Sub



Private Sub jnsf_Click()
Form21.Show
End Sub

Private Sub llyf_Click()
Form32.Show
End Sub

Private Sub mmxg_Click()
Form12.Show
End Sub

Private Sub TC_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(2).Text = "当前时间：" & Time()
End Sub
