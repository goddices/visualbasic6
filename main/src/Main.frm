VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form Mainfrm 
   Caption         =   "枫叶图书管理系统"
   ClientHeight    =   7980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   11880
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame6 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3855
      Left            =   120
      TabIndex        =   41
      Top             =   1320
      Width           =   615
      Begin VB.CommandButton CmdLogin 
         BackColor       =   &H000080FF&
         Caption         =   "登录(&D)"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton cmdKong 
         BackColor       =   &H000080FF&
         Caption         =   "清空(&K)"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -120
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton cmdBackBook 
         Caption         =   "还书(&H)"
         Height          =   255
         Left            =   -240
         TabIndex        =   42
         Top             =   4200
         Width           =   255
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   5805
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   10239
         ButtonWidth     =   1032
         ButtonHeight    =   1455
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList4"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "登录"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "清空"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "还书"
               ImageIndex      =   7
            EndProperty
         EndProperty
         BorderStyle     =   1
         Begin MSComctlLib.ImageList ImageList4 
            Left            =   120
            Top             =   3600
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   7
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":0442
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":0D1E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":0E7A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":12CE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":283A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":2C8E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":41E2
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "归还图书"
         ForeColor       =   &H00C000C0&
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   49
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ALT + D "
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   4
         Left            =   2520
         TabIndex        =   48
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ALT + K"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   5
         Left            =   2520
         TabIndex        =   47
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ALT + S"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   6
         Left            =   2520
         TabIndex        =   46
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ALT + C"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   7
         Left            =   2520
         TabIndex        =   45
         Top             =   1800
         Width           =   630
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ALT + H"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   9
         Left            =   2520
         TabIndex        =   44
         Top             =   720
         Width           =   720
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   945
      Left            =   0
      TabIndex        =   37
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1667
      BandCount       =   2
      _CBWidth        =   8175
      _CBHeight       =   945
      _Version        =   "6.0.8169"
      Child1          =   "Toolbar1"
      MinHeight1      =   885
      Width1          =   10770
      NewRow1         =   0   'False
      MinHeight2      =   360
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   885
         Left            =   165
         TabIndex        =   38
         Top             =   30
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   1561
         ButtonWidth     =   1879
         ButtonHeight    =   1455
         ImageList       =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "图书"
               Object.ToolTipText     =   "添加编辑图书"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "添加"
                     Text            =   "添加新书(&T)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "编辑"
                     Text            =   "编辑图书(&B)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "借书证"
               Object.ToolTipText     =   "添加编辑借书证"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "管理员"
               Object.ToolTipText     =   "修改管理员"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "图书分类"
               ImageIndex      =   11
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "设置"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               ImageIndex      =   2
            EndProperty
         EndProperty
         MousePointer    =   4
         Begin MSComctlLib.ImageList ImageList3 
            Left            =   6720
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   13
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":4636
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":4A8A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":4EDE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":5332
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":564E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":5AA2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":5EF6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":634A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":6666
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":6982
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":725E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":8A22
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Main.frx":8E76
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.StatusBar Sb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   36
      Top             =   7605
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5293
            MinWidth        =   5293
            Picture         =   "Main.frx":92CA
            Text            =   "程序设计语言:VB"
            TextSave        =   "程序设计语言:VB"
            Object.ToolTipText     =   "路亦超制作"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Picture         =   "Main.frx":971E
            Text            =   "软件作者:路亦超"
            TextSave        =   "软件作者:路亦超"
            Object.ToolTipText     =   "设计者"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Picture         =   "Main.frx":9B72
            TextSave        =   "23:02"
            Object.ToolTipText     =   "当前时间"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   3598
            MinWidth        =   3598
            Picture         =   "Main.frx":9FC6
            TextSave        =   "2011-12-21"
            Object.ToolTipText     =   "当前日期"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   4920
      TabIndex        =   32
      Top             =   2400
      Width           =   6975
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3360
         Top             =   480
      End
      Begin VB.Image Image3 
         Height          =   3000
         Left            =   0
         Picture         =   "Main.frx":A2E2
         Top             =   0
         Width           =   6930
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "欢迎使用枫叶图书管理系统！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1695
         Left            =   360
         TabIndex        =   40
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "借书"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4215
      Left            =   4920
      TabIndex        =   13
      Top             =   1200
      Width           =   6975
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   4800
         Top             =   2040
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Main.frx":E059
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Main.frx":E4AD
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Main.frx":E901
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Main.frx":ED55
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         Height          =   2775
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   6615
         Begin VB.CommandButton cmdOkCancel 
            BackColor       =   &H00C0C0FF&
            Caption         =   "借出图书(&S)"
            BeginProperty Font 
               Name            =   "楷体_GB2312"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txtType 
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
            Height          =   315
            Left            =   5280
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtLentDate 
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
            Height          =   345
            Left            =   1080
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txtChuBan 
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
            Height          =   300
            Left            =   1080
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   1560
            Width           =   3375
         End
         Begin VB.TextBox txtBookName 
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
            Height          =   330
            Left            =   1080
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   960
            Width           =   3375
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   3360
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtBookHao 
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
            Height          =   315
            Left            =   1080
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   5160
            Picture         =   "Main.frx":F1A9
            Top             =   1200
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "类别"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   10
            Left            =   4680
            TabIndex        =   29
            Top             =   480
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "今天日期"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   8
            Left            =   120
            TabIndex        =   24
            Top             =   2160
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "出版社"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   7
            Left            =   120
            TabIndex        =   19
            Top             =   1560
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "价格"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   6
            Left            =   2760
            TabIndex        =   18
            Top             =   480
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "书    名"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   5
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "图书编号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   4
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   840
         End
      End
      Begin VB.TextBox txtBookBian 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   2520
         TabIndex        =   1
         Top             =   540
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Enter"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5160
         TabIndex        =   39
         Top             =   840
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   5040
         Picture         =   "Main.frx":F5EB
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "输入要借的图书编号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2310
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000D&
         Index           =   0
         X1              =   120
         X2              =   7080
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "已借图书"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   11775
      Begin MSComctlLib.ListView LV2 
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   3201
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388736
         BackColor       =   12648447
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "仿宋_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "已经借出的图书"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4800
         TabIndex        =   5
         Top             =   240
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "基本信息"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4215
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
      Begin VB.TextBox txtFa 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   1200
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   3645
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame5"
         Height          =   135
         Left            =   3720
         TabIndex        =   28
         Top             =   0
         Width           =   15
      End
      Begin VB.TextBox txtZhiCheng 
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
         Height          =   330
         Left            =   1200
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   3000
         Width           =   1695
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5400
         Top             =   4200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Main.frx":FA2D
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Main.frx":FE81
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Main.frx":102D5
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtDepart 
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
         Height          =   345
         Left            =   1200
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtClass 
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
         Height          =   345
         Left            =   1200
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtName 
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
         Height          =   345
         Left            =   1200
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtBookId 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1200
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "元"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   12
         Left            =   2640
         TabIndex        =   35
         Top             =   3720
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "共计欠费"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   11
         Left            =   240
         TabIndex        =   33
         Top             =   3720
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "职   称"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   9
         Left            =   240
         TabIndex        =   26
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "部   门"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "班   级"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "姓   名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "借书证号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   840
      End
   End
   Begin VB.Line Line3 
      X1              =   11880
      X2              =   0
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Menu MnuFile 
      Caption         =   "控制(&C)"
      Begin VB.Menu FenMnu 
         Caption         =   "图书分类(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu s5 
         Caption         =   "-"
      End
      Begin VB.Menu AddMnu 
         Caption         =   "添加新书(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu EditMnu 
         Caption         =   "编辑图书(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu EditIdMnu 
         Caption         =   "编辑借书证(&J)"
         Shortcut        =   ^J
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMnu 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu ToolMnu 
      Caption         =   "功能(&T)"
      Begin VB.Menu LoginMnu 
         Caption         =   "登录(&D)"
      End
      Begin VB.Menu SearchMnu 
         Caption         =   "查找(&F)"
      End
      Begin VB.Menu BackMnu 
         Caption         =   "还书(&H)"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu SetMnu 
         Caption         =   "设置(&S)"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu HelpMnu 
      Caption         =   "帮助(&H)"
      Begin VB.Menu AboutMnu 
         Caption         =   "关于本软件(&A)"
      End
   End
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db1 As Database
Dim db2 As Database
Dim db3 As Database
Dim rst As Recordset
Dim rst1 As Recordset '打开表Personal
Dim rst2 As Recordset '打开表BookFlag
Dim rst3 As Recordset '打开表Book
Dim ws1 As Workspace
Dim ws2 As Workspace
Dim qry2 As QueryDef
Dim RecNumBookFf As Integer '表BookFf的记录个数
Dim SFlag As String
Private Type MSet
    BookNum As Integer
    BookCost As Single
End Type
Dim SetFlag As MSet
Option Explicit
Private Sub AboutMnu_Click()
Aboutfrm.Show (1)
End Sub



Private Sub AddMnu_Click()
Sb.Panels(1).Text = "添加新书"
        AddNewBook.Show (1)
        Sb.Panels(1).Text = SFlag
End Sub

Private Sub BackMnu_Click()
cmdBackBook_Click
End Sub

Private Sub cmdBackBook_Click() '打开还书对话框
cmdKong_Click
Lentfrm.Show (1)
cmdKong_Click
End Sub
Private Sub cmdKong_Click() '清空所有文本
txtBookId.Text = ""
txtName.Text = ""
txtClass.Text = ""
txtDepart.Text = ""
txtBookHao.Text = ""
txtBookName = ""
txtZhiCheng = ""
txtFa.Text = ""
txtBookBian.Text = ""
Frame4.Visible = False
Frame7.Visible = True
LV2.ListItems.Clear
CmdLogin.SetFocus
End Sub
Private Sub cmdOkCancel_Click(Index As Integer)
Select Case Index
    Case 1
        If rst3.Fields("是否借出") = True Then
            MsgBox "此书已经借出！", 0 + 48, "提示"
            txtBookBian.Text = ""
            txtBookBian.SetFocus
            Frame4.Visible = False
            Frame7.Visible = True
            Exit Sub
        End If
        rst2.AddNew
        rst2.Fields("图书编号") = rst3.Fields("图书编号")
        rst2.Fields("书名") = rst3.Fields("书名")
        rst2.Fields("价格") = rst3.Fields("价格")
        rst2.Fields("出版社") = rst3.Fields("出版社")
        rst2.Fields("借出日期") = rst3.Fields("借出日期")
        rst2.Fields("借书证号") = BookId
        rst2.Fields("姓名") = txtName.Text
        rst2.Fields("类别") = rst3.Fields("类别")
        rst2.Update
        rst3.Edit
        rst3.Fields("是否借出") = True
        rst3.Fields("借出日期") = rst3.Fields("借出日期")
        rst3.Update
        DataRef
        txtBookBian.Text = ""
        txtBookBian.SetFocus
        'CmdLogin.SetFocus
        Frame4.Visible = False
        Frame7.Visible = True
End Select
End Sub
Private Sub CmdLogin_Click()
loop1:  '如果没有此证，返回
LentLogin.Show (1)
If LoginFlag Then
LV2.ListItems.Clear
rst1.Seek "=", BookId  '查找借书证号码
If rst1.NoMatch Then
    MsgBox "没有此借书证号码！", 0 + 48, "错误"
    LoginFlag = False
    GoTo loop1  '返回loop1
End If
txtBookId.Text = BookId
txtName.Text = rst1.Fields("姓名") & vbNullString
txtClass.Text = rst1.Fields("班级") & vbNullString
txtDepart.Text = rst1.Fields("部门") & vbNullString
txtZhiCheng = rst1.Fields("职称") & vbNullString
txtFa.Text = rst1.Fields("罚款") & Empty
txtBookBian.Text = ""
Frame4.Visible = False
Frame7.Visible = True
txtBookBian.SetFocus
DataRef '输出所借图书
LoginFlag = False
If rst1.Fields("罚款") > 0 Then
   If MsgBox(txtBookId & " " & txtName & " 共计欠费 " _
        & rst1.Fields("罚款") & "元 是否从数据库中删除？", 4 + 48, "欠费") _
            = vbYes Then
        '从数据库中删除欠费记录
        rst1.Edit
        rst1.Fields("罚款") = 0
        rst1.Update
        txtFa.Text = rst1.Fields("罚款") & Empty
    End If
Else            '把罚款复制为0
    rst1.Edit
    rst1.Fields("罚款") = 0
    rst1.Update
End If

End If
End Sub





Private Sub EditIdMnu_Click()
Sb.Panels(1).Text = "编辑借书证"
        EditBookId.Show (1)
        Sb.Panels(1).Text = "编辑借书证"
End Sub

Private Sub EditMnu_Click()
        Sb.Panels(1).Text = "编辑图书"
        EditBook.Show (1)
        Sb.Panels(1).Text = "编辑图书"
End Sub

Private Sub ExitMnu_Click()
'Unload Me
End
End Sub

Private Sub FenMnu_Click()
SetType.Show (1)
End Sub

Private Sub Form_Load()
Set db1 = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst1 = db1.OpenRecordset("Personal", dbOpenTable)
rst1.Index = "借书证号"


Set db2 = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst2 = db2.OpenRecordset("BookFf", dbOpenTable)
Set qry2 = db2.CreateQueryDef("")
rst2.Index = "图书编号"

Set db3 = Workspaces(0).OpenDatabase(App.Path & "\DataBase\Data.mdb", False)
Set rst3 = db3.OpenRecordset("Book", dbOpenTable)
rst3.Index = "图书编号"

Open App.Path & "\DataBase\Data.mdb" For Random As #1 Len = Len(SetFlag)
Get #1, 1, SetFlag
BookNum = SetFlag.BookNum
FaCost = SetFlag.BookCost

LV2.View = lvwReport
LV2.ColumnHeaders.Add , , "借书证号"
LV2.ColumnHeaders.Add , , "借书人姓名"
LV2.ColumnHeaders.Add , , "图书编号"
LV2.ColumnHeaders.Add , , "书名"
LV2.ColumnHeaders.Add , , "价格"
LV2.ColumnHeaders.Add , , "类别"
LV2.ColumnHeaders.Add , , "出版社"
LV2.ColumnHeaders.Add , , "借出日期"

SFlag = "程序设计: 风云"

txtBookId.Text = ""
txtName.Text = ""
txtClass.Text = ""
txtDepart.Text = ""
txtBookHao.Text = ""
txtBookName = ""
txtZhiCheng = ""
txtFa.Text = ""

txtCost = ""
txtChuBan = ""
txtLentDate = ""


End Sub

Private Sub Form_Unload(Cancel As Integer)
rst1.Close
rst2.Close
rst3.Close
db1.Close
db2.Close
db3.Close
Close #1
End Sub


Private Sub LoginMnu_Click()
 CmdLogin_Click
End Sub



Private Sub SearchMnu_Click()
 Findfrm.Show
End Sub

Private Sub SetMnu_Click()
setfrm.Show
End Sub

Private Sub Timer1_Timer()

Me.Label4.Top = Me.Label4.Top + 10

If Me.Label4.Top >= Image3.Height Then

Me.Label4.Top = 0
End If

End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 3
        Sb.Panels(1).Text = "编辑借书证"
        EditBookId.Show (1)
        Sb.Panels(1).Text = "编辑借书证"
    Case 5
        SetPer.Show (1)
        Case 7
        SetType.Show
        Case 9
        setfrm.Show
        Case 13
        End
        
        
        
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal buttonmenu As MSComctlLib.buttonmenu)
Select Case buttonmenu.Key
    Case "添加"
        Sb.Panels(1).Text = "添加新书"
        AddNewBook.Show (1)
        Sb.Panels(1).Text = SFlag
    Case "编辑"
        Sb.Panels(1).Text = "编辑图书"
        EditBook.Show (1)
        Sb.Panels(1).Text = "编辑图书"
    Case "新借"
        MsgBox "Add BookCard"
    Case "编借"
        MsgBox "Edit BookCard"
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        CmdLogin_Click
    Case 3
        cmdKong_Click
    Case 7
        cmdBackBook_Click
        
    Case 5
        Findfrm.Show
End Select
End Sub


Private Sub txtBookBian_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtBookId.Text = "" Then
        MsgBox "请先登录！", 0 + 48, "提示"
        CmdLogin.SetFocus
        txtBookBian.Text = ""
        Exit Sub
    End If
    rst3.Seek "=", txtBookBian.Text
    If rst3.NoMatch Then
        MsgBox "没有此图书编号，请重新填写", 0 + 48, "填写错误"
        txtBookBian.SetFocus
        Frame4.Visible = False
        Frame7.Visible = True
        Exit Sub
    End If
    Frame4.Visible = True
    Frame7.Visible = False
    txtBookHao.Text = txtBookBian.Text
    txtBookName.Text = rst3.Fields("书名") & vbNullString
    txtChuBan.Text = rst3.Fields("出版社") & vbNullString
    txtCost.Text = rst3.Fields("价格") & Empty
    txtLentDate = rst3.Fields("借出日期") & vbNullString
    txtType.Text = rst3.Fields("类别") & vbNullString
End If
End Sub
Private Sub DataRef()
Dim i As Integer
Dim SeaStr As String
SeaStr = "select * from Bookff where 借书证号="
SeaStr = SeaStr & "'" & BookId & "'"
qry2.SQL = SeaStr
Set rst = qry2.OpenRecordset()
If rst.RecordCount = 0 Then
     Label1.Caption = "可以借" & BookNum & "本书"
     Exit Sub
End If
rst.MoveLast
RecNumBookFf = rst.RecordCount
rst.MoveFirst
LV2.ListItems.Clear
For i = 1 To RecNumBookFf
    LV2.ListItems.Add i, , rst.Fields("借书证号") & vbNullString
    With LV2.ListItems(i)
        .SubItems(1) = rst.Fields("姓名") & vbNullString
         .SubItems(2) = rst.Fields("图书编号") & vbNullString
        .SubItems(3) = rst.Fields("书名") & vbNullString
        .SubItems(4) = rst.Fields("价格") & Empty
        .SubItems(5) = rst.Fields("类别") & vbNullString
        .SubItems(6) = rst.Fields("出版社") & vbNullString
        .SubItems(7) = rst.Fields("借出日期") & vbNullString
    End With
    rst.MoveNext
    If rst.EOF Then Exit For
Next i
If RecNumBookFf = BookNum Then
    MsgBox "已经借了 " & BookNum & "本书,不能再借了,请登录其它借书证号", 0 + 48, "提示"
    txtBookId.Text = ""
    txtName.Text = ""
    txtClass.Text = ""
    txtDepart.Text = ""
    txtZhiCheng = ""
    txtFa.Text = ""
    CmdLogin.SetFocus
    LV2.ListItems.Clear
    Label1.Caption = "已经借的书"
    Exit Sub
End If
Label1.Caption = "已经借出 " & RecNumBookFf & "本，还可以再借 " _
        & BookNum - RecNumBookFf & "本"
End Sub
Private Sub txtBookId_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    LV2.ListItems.Clear
    BookId = txtBookId
    rst1.Seek "=", BookId  '查找借书证号码
    If rst1.NoMatch Then
     MsgBox "没有此借书证号码！", 0 + 48, "错误"
     txtBookId.SetFocus
     txtName.Text = ""
     txtClass.Text = ""
     txtDepart.Text = ""
     Exit Sub
    End If
        txtBookHao.Text = ""
        txtBookName.Text = ""
        txtCost.Text = ""
        txtChuBan.Text = ""
        txtLentDate.Text = ""
        txtBookBian.Text = ""
    txtBookId.Text = BookId
    txtName.Text = rst1.Fields("姓名") & vbNullString
    txtClass.Text = rst1.Fields("班级") & vbNullString
    txtDepart.Text = rst1.Fields("部门") & vbNullString
    txtZhiCheng = rst1.Fields("职称") & vbNullString
    txtFa.Text = rst1.Fields("罚款") & Empty
    txtBookBian.SetFocus
    DataRef '输出所借图书
End If
End Sub

 
