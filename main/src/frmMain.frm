VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{05589FA0-C356-11CE-BF01-00AA0055595A}#2.0#0"; "amcompat.tlb"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "LookIt"
   ClientHeight    =   6000
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7425
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   7425
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picMouseUp 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4725
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMouseDown 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4725
      Picture         =   "frmMain.frx":045C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   1260
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMouseOver 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4725
      Picture         =   "frmMain.frx":05AE
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   735
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   4695
      ScaleHeight     =   795
      ScaleWidth      =   960
      TabIndex        =   13
      Top             =   -810
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picSplitter2 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Height          =   60
      Left            =   30
      ScaleHeight     =   60
      ScaleWidth      =   2100
      TabIndex        =   12
      Top             =   3285
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.PictureBox pixSmall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   -240
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   10
      Top             =   390
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   -540
      Picture         =   "frmMain.frx":0700
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   -540
      Picture         =   "frmMain.frx":0802
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   105
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox pixLarger 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   -2700
      Picture         =   "frmMain.frx":0B0C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   15
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   15
      Top             =   5610
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E16
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   585
      Top             =   5610
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      FillColor       =   &H0000C000&
      FillStyle       =   6  'Cross
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   6
      Top             =   705
      Visible         =   0   'False
      Width           =   72
   End
   Begin VB.PictureBox picTitles 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2040
      ScaleHeight     =   300
      ScaleWidth      =   4395
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   420
      Width           =   4395
      Begin VB.TextBox DisplayPath 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   0
         Width           =   4170
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   5730
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2468
            MinWidth        =   2468
            Text            =   "����ͳ��"
            TextSave        =   "����ͳ��"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2116
            MinWidth        =   2116
            Text            =   "����ͳ��"
            TextSave        =   "����ͳ��"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   2822
            MinWidth        =   2822
            Text            =   "��ǰλ��"
            TextSave        =   "��ǰλ��"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "�ļ�����"
            TextSave        =   "�ļ�����"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��С"
            TextSave        =   "��С"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "�ļ�����"
            TextSave        =   "�ļ�����"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "��������"
            TextSave        =   "��������"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "10:45"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Text            =   "Yusilong"
            TextSave        =   "Yusilong"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   26
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "��"
            Object.ToolTipText     =   "��"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "����"
            Object.ToolTipText     =   "���������"
            ImageKey        =   "Clipboard"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "����"
            Object.ToolTipText     =   "���Ƶ�������"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ɾ��"
            Object.ToolTipText     =   "ɾ��"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Properties"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "��ӡ"
            Object.ToolTipText     =   "��ӡ"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "��ͼ��"
            Object.ToolTipText     =   "��ͼ��"
            ImageKey        =   "View Large Icons"
            Style           =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Сͼ��"
            Object.ToolTipText     =   "Сͼ��"
            ImageKey        =   "View Small Icons"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "�б�"
            Object.ToolTipText     =   "�б�"
            ImageKey        =   "View List"
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "��ϸ����"
            Object.ToolTipText     =   "��ϸ����"
            ImageKey        =   "View Details"
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Sort Ascending"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Sort Descending"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "������"
            Object.ToolTipText     =   "������"
            ImageKey        =   "Tools"
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "״̬��"
            Object.ToolTipText     =   "״̬��"
            ImageKey        =   "Status"
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ԥ����"
            Object.ToolTipText     =   "Ԥ����"
            ImageKey        =   "Preview"
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   4800
      Left            =   2055
      TabIndex        =   1
      Top             =   690
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   8467
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "��С"
         Object.Width           =   2471
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "�޸�ʱ��"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "����"
         Object.Width           =   1058
      EndProperty
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   15
      ScaleHeight     =   1080
      ScaleWidth      =   1965
      TabIndex        =   11
      Top             =   4470
      Width           =   1965
      Begin SHDocVwCtl.WebBrowser GifView 
         Height          =   1050
         Left            =   210
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   15
         Visible         =   0   'False
         Width           =   165
         ExtentX         =   291
         ExtentY         =   1852
         ViewMode        =   1
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   -1  'True
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.PictureBox picShow 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   1050
         Left            =   -435
         MousePointer    =   99  'Custom
         ScaleHeight     =   1050
         ScaleWidth      =   630
         TabIndex        =   14
         Top             =   15
         Visible         =   0   'False
         Width           =   630
      End
      Begin AMovieCtl.ActiveMovie AudioDisplay 
         Height          =   1050
         Left            =   45
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   15
         Visible         =   0   'False
         Width           =   405
         Appearance      =   0
         AutoStart       =   0   'False
         AllowChangeDisplayMode=   -1  'True
         AllowHideDisplay=   0   'False
         AllowHideControls=   -1  'True
         AutoRewind      =   0   'False
         Balance         =   0
         CurrentPosition =   0
         DisplayBackColor=   192
         DisplayForeColor=   65535
         DisplayMode     =   0
         Enabled         =   -1  'True
         EnableContextMenu=   0   'False
         EnablePositionControls=   -1  'True
         EnableSelectionControls=   0   'False
         EnableTracker   =   -1  'True
         Filename        =   ""
         FullScreenMode  =   0   'False
         MovieWindowSize =   0
         PlayCount       =   1
         Rate            =   1
         SelectionStart  =   -1
         SelectionEnd    =   -1
         ShowControls    =   -1  'True
         ShowDisplay     =   -1  'True
         ShowPositionControls=   -1  'True
         ShowTracker     =   -1  'True
         Volume          =   -60
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   15
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":184E
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1960
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A72
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B84
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C96
            Key             =   "Clipboard"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DA8
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EBA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FCC
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20DE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21F0
            Key             =   "Tools"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2302
            Key             =   "Status"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2414
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2526
            Key             =   "Sort Descending"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2638
            Key             =   "Sort Ascending"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":274A
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":285C
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":296E
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A80
            Key             =   "View Details"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B92
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin LookIt.vbwFolderView tvTreeView 
      Height          =   2370
      Left            =   60
      TabIndex        =   19
      Top             =   495
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4180
   End
   Begin VB.Image imgSplitter 
      Height          =   4755
      Left            =   1860
      MouseIcon       =   "frmMain.frx":2CA4
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":2DF6
      Stretch         =   -1  'True
      ToolTipText     =   "�϶�"
      Top             =   -390
      Width           =   90
   End
   Begin VB.Image imgSplitter2 
      Height          =   120
      Left            =   -195
      MouseIcon       =   "frmMain.frx":46FC
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":484E
      Stretch         =   -1  'True
      ToolTipText     =   "�϶�"
      Top             =   3075
      Width           =   4140
   End
   Begin VB.Menu MnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "���ļ�(&O)"
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuFile002 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrinterSet 
         Caption         =   "��ӡ����(&S) ..."
      End
      Begin VB.Menu MnuPrintPicture 
         Caption         =   "��ӡ(&P) ..."
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuLine001 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "ѡ��(&O)..."
      End
      Begin VB.Menu MnuFile001 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "�ر�(&C)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuVideo 
      Caption         =   "Ӱ��(&V)"
      Visible         =   0   'False
      Begin VB.Menu MnuVideoPlay 
         Caption         =   "����(&P)"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu MnuVideoPause 
         Caption         =   "��ͣ(&U)"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu MnuVideoStop 
         Caption         =   "ֹͣ(&S)"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu MnuVideoLine01 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFullScreen 
         Caption         =   "ȫ����ʾ(&F)"
         Checked         =   -1  'True
         Shortcut        =   +^{F8}
      End
      Begin VB.Menu MnuVideoLine02 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMediaPlay 
         Caption         =   "����VCD������(M) ..."
      End
   End
   Begin VB.Menu MnuPicture 
      Caption         =   "ͼƬ(&P)"
      Visible         =   0   'False
      Begin VB.Menu MnuPictureView 
         Caption         =   "�鿴(&V)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuShowSize 
         Caption         =   "�Զ���С(&A)"
         Index           =   0
      End
      Begin VB.Menu MnuShowSize 
         Caption         =   "������С"
         Index           =   1
      End
      Begin VB.Menu MnuShowSize 
         Caption         =   "1/2 ��С"
         Index           =   2
      End
      Begin VB.Menu MnuShowSize 
         Caption         =   "1/4 ��С"
         Index           =   3
      End
      Begin VB.Menu MnuShowSize 
         Caption         =   "1/8 ��С"
         Index           =   4
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSetBackGroundM 
         Caption         =   "����Ϊǽֽ"
         Begin VB.Menu MnuSetBackground 
            Caption         =   "��ǽֽ����(&C)    "
            Index           =   0
         End
         Begin VB.Menu MnuSetBackground 
            Caption         =   "��ǽֽ����(&E)"
            Index           =   1
         End
         Begin VB.Menu MnuSetBackground 
            Caption         =   "��ǽֽƽ��(&T)"
            Index           =   2
         End
         Begin VB.Menu mnuFileBar2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCleanBackground 
            Caption         =   "�������ǽֽ"
            Shortcut        =   ^E
         End
      End
   End
   Begin VB.Menu mnuMainView 
      Caption         =   "�༭(&E)"
      Begin VB.Menu MnuLookFor 
         Caption         =   "�鿴ͼƬ(&L) ..."
         Enabled         =   0   'False
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFileOpenAs 
         Caption         =   "�򿪷�ʽ(&E) ..."
         Visible         =   0   'False
      End
      Begin VB.Menu MnuOpenFileAsLine 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDefineFace 
         Caption         =   "���涨��(&I)"
         Begin VB.Menu mnuViewToolbar 
            Caption         =   "������(&T)   "
            Checked         =   -1  'True
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuViewStatusBar 
            Caption         =   "״̬��(&B)   "
            Checked         =   -1  'True
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu MnuViewPreview 
            Caption         =   "Ԥ����(&V)   "
            Checked         =   -1  'True
            Shortcut        =   ^{F3}
         End
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuViewMain 
         Caption         =   "�鿴ͼ��(&V)"
         Begin VB.Menu MnuView 
            Caption         =   "��ʾ��ͼ��(&G)  "
            Index           =   0
            Shortcut        =   +{F1}
         End
         Begin VB.Menu MnuView 
            Caption         =   "��ʾСͼ��(&M)  "
            Index           =   1
            Shortcut        =   +{F2}
         End
         Begin VB.Menu MnuView 
            Caption         =   "��ʾ�б�(&L)"
            Index           =   2
            Shortcut        =   +{F3}
         End
         Begin VB.Menu MnuView 
            Caption         =   "��ʾ��ϸ����(&D)  "
            Index           =   3
            Shortcut        =   +{F4}
         End
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArrangeIcon 
         Caption         =   "����ͼ��(&A)"
         Begin VB.Menu MnuArrangSort 
            Caption         =   "����������(&N)   "
            Index           =   0
            Shortcut        =   +{F5}
         End
         Begin VB.Menu MnuArrangSort 
            Caption         =   "����С����(&S)"
            Index           =   1
            Shortcut        =   +{F6}
         End
         Begin VB.Menu MnuArrangSort 
            Caption         =   "����������(&T)"
            Index           =   2
            Shortcut        =   +{F7}
         End
         Begin VB.Menu MnuArrangSort 
            Caption         =   "����������(&D)"
            Index           =   3
            Shortcut        =   +{F8}
         End
         Begin VB.Menu mnuFileBar5 
            Caption         =   "-"
         End
         Begin VB.Menu MnuArrangSortAuto 
            Caption         =   "����������(&A)"
            Shortcut        =   +{F11}
         End
         Begin VB.Menu MnuArrangSortAutoZ 
            Caption         =   "����������(&E)"
            Shortcut        =   +{F12}
         End
         Begin VB.Menu Line0002 
            Caption         =   "-"
         End
         Begin VB.Menu mnuArrangeFileIcon 
            Caption         =   "�Զ�����ͼ��(&U)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "����(&C)"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuLine0002 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCopyTo 
         Caption         =   "���Ƶ�(&T)..."
         Enabled         =   0   'False
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuEditMove 
         Caption         =   "�ƶ���(&M)..."
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuClearClipboard 
         Caption         =   "��� Clipboard"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "ɾ��(&D)"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "������(&N)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRefreshDir 
         Caption         =   "ˢ��Ŀ¼(&D)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ���б�(&F)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MnuLine0003 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileAttribute 
         Caption         =   "����(&R)"
         Shortcut        =   ^{F12}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "����(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuDisplayPictureViewWindow 
         Caption         =   "ͼƬ�鿴����(&V) ..."
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMemdiaPlay 
         Caption         =   "����VCD������ (&M) ..."
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "���ı������ (&W) ..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Ŀ¼(&C)"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "������������(&S)..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A) "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'���幤������ʼ��
Const View_Number = 14
Const Display_Number = 22
Const Printer_Number = 12
Const Copy_Number = 7
Const sglSplitLimit = 50
Const PD_PRINTSETUP = &H40

Dim OldShowSize As Integer  '��ʾ��С
Dim mbMoving As Boolean, UndoK As Boolean, DisplayTrue As Boolean
Dim mlNextClipboardViewer As Long
Dim OldName As String
Dim OldItem As String, NewItem As String

'����Դ�ļ���Ŀ���ļ�
Public SourceFile As String
Public TargetFile As String

Private Type SHELLEXECUTEINFO
  cbSize As Long
  fMask As Long
  hWnd As Long
  lpVerb As String
  lpFile As String
  lpParameters As String
  lpDirectory As String
  nShow As Long
  hInstApp As Long
  lpIDList As Long
  lpClass As String
  hkeyClass As Long
  dwHotKey As Long
  hIcon As Long
  hProcess As Long
End Type

Private Declare Function ShellExecuteEx Lib "shell32" (lpSEI As SHELLEXECUTEINFO) As Long
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Const SEE_MASK_INVOKEIDLIST = &HC

Private Sub AudioDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  If Button = vbRightButton Then
     PopupMenu MnuVideo
  End If
   
End Sub

Private Sub AudioDisplay_OpenComplete()

  AudioDisplay.Visible = True
  
  If AudioDisplay.Width >= picDisplay.Width Then
     AudioDisplay.Left = 0
   Else
     AudioDisplay.Left = (picDisplay.Width - AudioDisplay.Width) / 2
  End If
  
  If AudioDisplay.Height >= picDisplay.Height Then
     AudioDisplay.Top = 0
   Else
     AudioDisplay.Top = (picDisplay.Height - AudioDisplay.Height) / 2
  End If
  
  lvListView.SetFocus
  
End Sub

Private Sub AudioDisplay_StateChange(ByVal oldState As Long, ByVal newState As Long)

  If AudioDisplay.CurrentState = amvRunning Then  '����ʱ��Ч
      tbToolBar.Buttons(4).Enabled = False
      MnuVideoPlay.Enabled = False
      MnuVideoPause.Enabled = True
      MnuVideoStop.Enabled = True
     Else
      tbToolBar.Buttons(4).Enabled = True
      MnuVideoPlay.Enabled = True
      MnuVideoPause.Enabled = False
      MnuVideoStop.Enabled = False
  End If
  
End Sub

Private Sub Form_Activate()
    
    If Not tvTreeView.bLoaded Then
           tvTreeView.Init
    End If
    'fPath$ = "C:\"  '������
    'vbGetFileList
    mlNextClipboardViewer = SetClipboardViewer(Me.hWnd)
       
   If DisplayTrue = False Then
      Call mnuView_Click(Val(GetSetting(App.Title, "Settings", "ViewMode", 0)))
      Call mnuViewRefresh_Click
      DisplayTrue = True
   End If
    
End Sub

Private Sub Form_Load()
    
    '��װ����
    
    IniData '��ʼ������
    
    picDisplay.Left = tvTreeView.Left
    imgSplitter2.Left = picDisplay.Left
    imgSplitter2.Width = imgSplitter.Left
    
    SubClass Me
    Me.Show
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    UnSubClass Me
    Call ChangeClipboardChain(Me.hWnd, mlNextClipboardViewer)
    'Dim i As Integer  'ж�������Ӵ���
    'For i = Forms.Count - 1 To 1 Step -1
    '    Unload Forms(i)
    'Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    SaveSetting App.Title, "Settings", "ViewMode", lvListView.View
    SaveSetting App.Title, "Settings", "Position", imgSplitter.Left
    SaveSetting App.Title, "Settings", "HPosition", imgSplitter2.Top
       
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub  '��С��ʱ�˳�
    
    If Me.Width < 5000 Then Me.Width = 5000
        
    SizeControls imgSplitter2.Width
    SizeControlsH imgSplitter2.Top
    
End Sub

Private Sub GifView_NavigateComplete2(ByVal pDisp As Object, URL As Variant)

  lvListView.SetFocus
 
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
    
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim sglPos As Single
    
    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
    
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
    
End Sub

Sub SizeControls(X As Single)

    On Error Resume Next

    '���� Width ����
    If X < 2500 Then X = 2500
    If X > (Me.Width - 2500) Then X = Me.Width - 2500
    tvTreeView.Width = X
    imgSplitter2.Width = X  '��ֱ��
    picDisplay.Width = X  'Ԥ����
    imgSplitter.Left = X + 60
    lvListView.Left = X + 150
    lvListView.Width = Me.Width - (tvTreeView.Width + 280)
    picTitles.Left = lvListView.Left
    picTitles.Width = lvListView.Width

    If tbToolBar.Visible Then
        tvTreeView.Top = tbToolBar.Height
        picTitles.Top = tbToolBar.Height
    Else
        tvTreeView.Top = 0
        picTitles.Top = 0
    End If

    lvListView.Top = tvTreeView.Top + picTitles.Height
       
    If sbStatusBar.Visible Then
       lvListView.Height = Me.ScaleHeight - (picTitles.Height + picTitles.Top) - sbStatusBar.Height
    Else
       lvListView.Height = Me.ScaleHeight - (picTitles.Height + picTitles.Top)
    End If
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = lvListView.Height + picTitles.Height
    
    DisplayPath.Width = lvListView.Width
    
End Sub

Private Sub imgSplitter2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    With imgSplitter2
        picSplitter2.Move .Left, .Top, .Width, .Height \ 2
    End With
      
    picSplitter2.Visible = True
    mbMoving = True
   
End Sub

Private Sub imgSplitter2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    Dim sglPosL As Single
    
    If mbMoving Then
       sglPosL = Y + imgSplitter2.Top
       If sglPosL < sglSplitLimit Then
          picSplitter2.Top = sglSplitLimit
       ElseIf sglPosL > Me.Height - sglSplitLimit Then
           picSplitter2.Top = Me.Height - sglSplitLimit
        Else
          picSplitter2.Top = sglPosL
       End If
    End If
    
    
End Sub

Private Sub imgSplitter2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    SizeControlsH picSplitter2.Top
    picSplitter2.Visible = False
    mbMoving = False
    
End Sub

Private Sub lvListView_AfterLabelEdit(Cancel As Integer, NewString As String)
 
  If Trim(NewString) = "" Then
     MsgBox "�Բ���,�ļ����Ʋ���Ϊ�ա�    ", vbCritical + vbOKOnly, "����������..."
     Cancel = -1                           'ȡ��������
     Exit Sub
  End If

  If Trim(NewString) = Trim(OldName) Then Exit Sub  '���ļ�������ļ�����ͬʱ
  
  '����ļ�����
   Dim SHop As SHFILEOPSTRUCT
   Dim strFile As String
   strFile = ValidateDir(fPath$) & OldName
   
   With SHop
      .wFunc = FO_RENAME
      .pFrom = strFile
      .pTo = ValidateDir(fPath$) & NewString
      .fFlags = FOF_NOCONFIRMATION
   End With
   
   Dim retVal As Long   'ִ��
   retVal = SHFileOperation(SHop)
   
   If retVal <> 0 Then  '����ִ��ʱȡ������
      Cancel = -1
   End If

End Sub

Private Sub lvListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

  lvListView.SortKey = ColumnHeader.Index - 1
  lvListView.SortOrder = lvwAscending
  lvListView.Sorted = True
  
End Sub

Private Sub lvListView_DblClick()

 If lvListView.ListItems.Count > 0 Then
   If lvListView.SelectedItem.Selected Then
      OpenSelected
   End If
 End If
 
End Sub

Private Sub lvListView_ItemClick(ByVal Item As MSComctlLib.ListItem)

   '����Ƿ�Ϊͬһ��Ŀ
   NewItem = Item.Text
   If NewItem = OldItem Then Exit Sub
      OldItem = NewItem

   If Item.Text > "" Then
       MenuEnabled (-1)
   Else
       MenuEnabled (0)
   End If
     
   fPath$ = ValidateDir(fPath$)
   Dim picFile As String
       picFile = fPath$ & Item.Text
       SourceFile = picFile
'�����ļ�����
 Select Case VbGetFileType(Item.Text)
  Case "ͼƬ"
     '����ͼƬ
     PictureProccess (picFile)
  Case "����"
     GifProccess (picFile)
  Case "�ı�"
     txtProccess (picFile)
  Case "����"
     AudioProccess (picFile)
     SourceFile = ""  '����������ý�岥���豸
  Case "Ӱ��"
     VideoProccess (picFile)
  Case Else
   If mnuEditCopy.Enabled = True Then  '�ϴ�ΪͼƬʱ
     mnuEditCopy.Enabled = False '���ư�ť��Ч
     MnuLookFor.Enabled = False  '�鿴�˵���Ч
     MnuPrintPicture.Enabled = False
    '����ͼƬ�˵�
     tbToolBar.Buttons(Copy_Number).Enabled = False
     tbToolBar.Buttons(Printer_Number).Enabled = False
     sbStatusBar.Panels(3).Text = "δע���: V1.0"
   End If
   If MnuVideo.Visible Then  '��Ƶ�˵�
      MnuVideo.Visible = False
   End If
   
  End Select

   sbStatusBar.Panels(4).Text = Item.Text
   sbStatusBar.Panels(5).Text = Item.ListSubItems(1).Text
   sbStatusBar.Panels(6).Text = Item.ListSubItems(2).Text
   sbStatusBar.Panels(7).Text = Item.ListSubItems(3).Text
   sbStatusBar.Panels(8).Text = Item.ListSubItems(4).Text
     
End Sub

Private Sub lvListView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
 If Shift = 1 Then
    UndoK = True
    mnuFileOpenAs.Visible = True
  Else
    UndoK = False
    mnuFileOpenAs.Visible = False
 End If
 
 If lvListView.ListItems.Count > 0 Then
    If lvListView.SelectedItem.Selected Then
       MenuEnabled (-1)
     Else
       MenuEnabled (0)
    End If
 End If
 
 If Button = 2 Then PopupMenu mnuMainView
   
End Sub

Private Sub mnuArrangeFileIcon_Click()

   mnuArrangeFileIcon.Checked = Not mnuArrangeFileIcon.Checked
   
   If mnuArrangeFileIcon.Checked = True Then
      lvListView.Arrange = lvwAutoTop
      SaveSetting App.Title, "Settings", "AutoArrange", 1
    Else
      lvListView.Arrange = lvwNone
      SaveSetting App.Title, "Settings", "AutoArrange", 0
   End If
   
   '����״̬
   
End Sub

Private Sub MnuArrangSort_Click(Index As Integer)
    
    lvListView.SortKey = Index
    lvListView.SortOrder = 0
    lvListView.Sorted = True
    
End Sub

Private Sub MnuArrangSortAuto_Click()
   
      lvListView.SortOrder = 0
      lvListView.Sorted = True
    
End Sub

Private Sub MnuArrangSortAutoZ_Click()
      
      lvListView.SortOrder = 1
      lvListView.Sorted = True

End Sub

Private Sub MnuCleanBackground_Click()

ChangePaper picBuffer, False
 
End Sub

Private Sub MnuClearClipboard_Click()
  
  Clipboard.Clear
  
End Sub

Private Sub mnuDisplayPictureViewWindow_Click()
  
  Screen.MousePointer = vbHourglass
  
  If Not frmPictureView.Visible Then
     Load frmPictureView
  End If
  
  If picLoad Then
     frmPictureView.picView.Picture = picBuffer.Picture
  End If
  
  Screen.MousePointer = vbDefault
  frmPictureView.Show vbNormal
  
End Sub

Private Sub mnuEditCopyTo_Click()

  TargetFile = SelectFilePath(Me.hWnd, "��ѡ���Ƶ���Ŀ¼��")

  If Trim(TargetFile) = "" Then  '������ڿ�ʱ�˳�
     Exit Sub
  End If
 
  TargetFile = ValidateDir(TargetFile) & lvListView.SelectedItem.Text
  SourceFile = ValidateDir(fPath$) & lvListView.SelectedItem.Text
     
  If SourceFile = TargetFile Then
     Exit Sub
  End If
  
  'ϵͳ��Shell�����ļ�
  Dim Result As Long, fileOp As SHFILEOPSTRUCT
   With fileOp
        .hWnd = Me.hWnd
        .wFunc = FO_COPY
        .pFrom = SourceFile
        .pTo = TargetFile
        .fFlags = FOF_SIMPLEPROGRESS + FOF_FILESONLY
   End With
   
   Result = SHFileOperation(fileOp)

End Sub

Private Sub mnuEditMove_Click()

TargetFile = SelectFilePath(Me.hWnd, "��ѡ���ƶ�����Ŀ¼��")

If Trim(TargetFile) = "" Then  '������ڿ�ʱ�˳�
   Exit Sub
End If
 
  TargetFile = ValidateDir(TargetFile) & lvListView.SelectedItem.Text
  SourceFile = ValidateDir(fPath$) & lvListView.SelectedItem.Text
  
  If SourceFile = TargetFile Then
     Exit Sub
  End If
     
  'ʹ��Name�����ƶ��ļ�
  'Name SourceFile As TargetFile
  'ϵͳ��Shell�ƶ��ļ�
  Dim Result As Long, fileOp As SHFILEOPSTRUCT
   With fileOp
        .hWnd = Me.hWnd
        .wFunc = FO_MOVE
        .pFrom = SourceFile
        .pTo = TargetFile
        .fFlags = FOF_SIMPLEPROGRESS + FOF_FILESONLY
   End With
   
   Result = SHFileOperation(fileOp)
    
  'ϵͳɾ���������ʱ,������û��ɾ��
   Result = GetFileAttributes(SourceFile)

  If Result = -1 Then '���ʱ
     lvListView.ListItems.Remove lvListView.SelectedItem.Index
     RefreshDesk
   Else
     Exit Sub 'û��ʱ
  End If
   
    
End Sub

Private Sub MnuFileAttribute_Click()

  '��ʾ�ļ�����
  If lvListView.ListItems.Count > 0 Then
     If lvListView.SelectedItem.Selected Then
        ShowFileProperties ValidateDir(fPath$) & lvListView.SelectedItem.Text
     End If
   ElseIf Trim(DisplayPath.Text) > "" Then
        ShowFileProperties Trim(DisplayPath.Text)
  End If
  
End Sub


Private Sub mnuFileOpenAs_Click()
 
 If lvListView.ListItems.Count > 0 Then
   If lvListView.SelectedItem.Selected Then
      OpenSelectedAs
   End If
 End If

End Sub

Private Sub mnuFileRename_Click()

   '�ļ�������
   If lvListView.ListItems.Count > 0 Then
      If lvListView.SelectedItem.Selected Then
      
         OldName = lvListView.SelectedItem.Text 'ȡ�ñ༭ǰ���ļ���
         lvListView.StartLabelEdit              '��ʼ�༭
         
      End If
   End If
   
End Sub

Private Sub MnuFullScreen_Click()

  MnuFullScreen.Checked = Not MnuFullScreen.Checked
  
  AudioDisplay.Pause
  
  If MnuFullScreen.Checked = True Then
     AudioDisplay.FullScreenMode = True
   Else
     AudioDisplay.FullScreenMode = False
  End If
  
  AudioDisplay.Run
  
End Sub

Private Sub MnuLookFor_Click()

  If picLoad = False Then  '��Ԥ������û��װ��ʱ
     Dim picFile As String
         picFile = fPath$ & lvListView.SelectedItem.Text
     Screen.MousePointer = vbHourglass
     picBuffer.Picture = LoadPicture(picFile)
     Screen.MousePointer = vbDefault
     picLoad = True    '�Ѿ���װ
  End If
    MnuPictureView_Click
 
End Sub

Private Sub MnuMediaPlay_Click()

 If AudioDisplay.Visible Then
    If AudioDisplay.CurrentState = amvRunning Then
       If SourceFile <> "" Then
          AudioDisplay.Stop
       End If
    End If
 End If
 
 Dim retVal As Long
     retVal = Shell("FlVcd3.0.Exe " & SourceFile, vbNormalFocus)
 If retVal = 0 Then
    MsgBox "�Բ���δ֪������������ý�岥����"
 End If
 
End Sub

Private Sub MnuMemdiaPlay_Click()
 
 If AudioDisplay.Visible Then
    If AudioDisplay.CurrentState = amvRunning Then
       If SourceFile <> "" Then
          AudioDisplay.Stop
       End If
    End If
 End If
  
 Dim retVal As Long
     retVal = Shell("FlVcd3.0.Exe " & SourceFile, vbNormalFocus)
 If retVal = 0 Then
    MsgBox "�Բ���δ֪������������ý�岥����"
 End If
 
End Sub

Private Sub MnuPictureView_Click()
  
  Screen.MousePointer = vbHourglass
  
  If frmPictureView.Visible Then
     frmPictureView.picView.Picture = picBuffer.Picture
   Else
     Load frmPictureView
     frmPictureView.picView.Picture = picBuffer.Picture
  End If
  
  Screen.MousePointer = vbDefault
  frmPictureView.Show vbNormal
  
End Sub

Private Sub MnuPrinterSet_Click()

  Dim setPrinter As New cCommonDialog
  
  setPrinter.CancelError = True
  setPrinter.flags = PD_PRINTSETUP
  
  setPrinter.ShowPrinter
  
     
End Sub

Private Sub MnuPrintPicture_Click()

   If picLoad = False Then
     Dim picFile As String
         picFile = fPath$ & lvListView.SelectedItem.Text
     Screen.MousePointer = vbHourglass
     picBuffer.Picture = LoadPicture(picFile)
     Screen.MousePointer = vbDefault
     picLoad = True '�Ѿ���װ���
   End If
  '��ʾ��ӡѡ��
   frmPicturePrint.Show 1
   
End Sub



Private Sub MnuRefreshDir_Click()

  If tvTreeView.bLoaded Then
     tvTreeView.UnInit
     tvTreeView.Init
     tvTreeView_SelectionChange "", ""
  End If
  
End Sub

Private Sub MnuSetBackground_Click(Index As Integer)
  
  Screen.MousePointer = vbHourglass
  '������ֲ�������
  Dim sKeyName As String, sEntry As String
  Dim sValue As String, bSuccess As Boolean
 
  Select Case Index
  Case 0
  '����ʱ
  sKeyName = "HKEY_CURRENT_USER\Control Panel\Desktop"
  sEntry = "WallpaperStyle"
  sValue = "0"
  bSuccess = WriteRegStringValue(sKeyName, sEntry, sValue)
  sKeyName = "HKEY_CURRENT_USER\Control Panel\Desktop"
  sEntry = "TileWallpaper"
  sValue = "0"
  bSuccess = WriteRegStringValue(sKeyName, sEntry, sValue)
 
  Case 1
  '��չʱ
  sKeyName = "HKEY_CURRENT_USER\Control Panel\Desktop"
  sEntry = "WallpaperStyle"
  sValue = "2"
  bSuccess = WriteRegStringValue(sKeyName, sEntry, sValue)
  sKeyName = "HKEY_CURRENT_USER\Control Panel\Desktop"
  sEntry = "TileWallpaper"
  sValue = "0"
  bSuccess = WriteRegStringValue(sKeyName, sEntry, sValue)
 
  Case 2
  'ƽ��ʱ
  sKeyName = "HKEY_CURRENT_USER\Control Panel\Desktop"
  sEntry = "WallpaperStyle"
  sValue = "1"
  bSuccess = WriteRegStringValue(sKeyName, sEntry, sValue)
  sKeyName = "HKEY_CURRENT_USER\Control Panel\Desktop"
  sEntry = "TileWallpaper"
  sValue = "1"
  bSuccess = WriteRegStringValue(sKeyName, sEntry, sValue)
  End Select
 
  ChangePaper picBuffer, True
  Screen.MousePointer = vbDefault

End Sub

Private Sub MnuShowSize_Click(Index As Integer)

  MnuShowSize(OldShowSize).Checked = False
  MnuShowSize(Index).Checked = True
  OldShowSize = Index
  SaveSetting App.Title, "Settings", "ShowSize", Index
    
  ShowPreview picBuffer, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picShow, picDisplay.ScaleWidth, picDisplay.ScaleHeight, picDisplay.Visible, OldShowSize
  
End Sub

Private Sub mnuToolsOptions_Click()

  frmOption.Show 1
  
End Sub

Private Sub MnuVideoPause_Click()
   
   AudioDisplay.Pause
   
End Sub

Private Sub MnuVideoPlay_Click()
 
   AudioDisplay.Run
   
End Sub

Private Sub MnuVideoStop_Click()
  
   AudioDisplay.Stop
  
End Sub

Private Sub mnuView_Click(Index As Integer)

   MnuView(lvListView.View).Checked = False  'ȡ���ϴεĲ鿴
   MnuView(Index).Checked = True   'ȷ���˴β鿴
    
   tbToolBar.Buttons(View_Number + Index).Value = tbrPressed
   
   lvListView.View = Index
   SaveSetting App.Title, "Settings", "ViewMode", Index

End Sub


Private Sub MnuViewPreview_Click()
    
    MnuViewPreview.Checked = Not MnuViewPreview.Checked
    
    'Ԥ�����ı�
    If MnuViewPreview.Checked = True Then
       tbToolBar.Buttons(Display_Number + 2).Value = tbrPressed
       SaveSetting App.Title, "Settings", "DisplayPreview", 1
    Else
       tbToolBar.Buttons(Display_Number + 2).Value = tbrUnpressed
       MnuPicture.Visible = False
       SaveSetting App.Title, "Settings", "DisplayPreview", 0
       RefreshDesk   'ˢ�°���
    End If
    picDisplay.Visible = MnuViewPreview.Checked
   
    SizeControlsH Val(GetSetting(App.Title, "Settings", "HPosition", 1500))
    SaveSetting App.Title, "Settings", "HPosition", imgSplitter2.Top
        
    'If picLoad = True Then
    '   If picDisplay.Visible Then
    '      Call picDisplay_Resize
    '   End If
    'End If
    
End Sub

Private Sub picDisplay_DblClick()

 'MnuPictureView_Click  '��ʾ�鿴
  
End Sub

Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
 'If Button = 2 And picShow.Visible = True Then PopupMenu MnuPicture

End Sub


Private Sub picDisplay_Resize()

If picShow.Visible Then  '���ͼƬ���ʱ
  If picShow.Height > picDisplay.Height Then
     picShow.Top = 0
     picShow.MouseIcon = picMouseOver.Picture
    Else
     picShow.Top = (picDisplay.Height - picShow.Height) / 2
  End If
  If picShow.Width > picDisplay.Width Then
     picShow.Left = 0
     picShow.MouseIcon = picMouseOver.Picture
    Else
     picShow.Left = (picDisplay.Width - picShow.Width) / 2
  End If
  '��װ���
  If picShow.ScaleHeight <= picDisplay.Height And picShow.ScaleWidth <= picDisplay.Width Then
     picShow.MouseIcon = picMouseUp.Picture
  End If
    
  'Ԥ������ͼƬʱ
  If picLoad = True And OldShowSize = 0 Then
     Screen.MousePointer = vbArrowHourglass
      '�Ƿ�װͼƬ
      If picDisplay.Visible Then
         ShowPreview picBuffer, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picShow, picDisplay.ScaleWidth, picDisplay.ScaleHeight, picDisplay.Visible, OldShowSize
      End If
     Screen.MousePointer = vbDefault
  End If
End If

If GifView.Visible Then  '���GIF��Чʱ
   GifView.Left = 20
   GifView.Height = picDisplay.ScaleHeight - 40
   GifView.Width = picDisplay.ScaleWidth - 40
End If

If AudioDisplay.Visible Then  '���������Чʱ
  If AudioDisplay.Width >= picDisplay.Width Then
     AudioDisplay.Left = 0
   Else
     AudioDisplay.Left = (picDisplay.Width - AudioDisplay.Width) / 2
  End If
  
  If AudioDisplay.Height >= picDisplay.Height Then
     AudioDisplay.Top = 0
   Else
     AudioDisplay.Top = (picDisplay.Height - AudioDisplay.Height) / 2
  End If
End If

End Sub

Private Sub picShow_DblClick()

  MnuPictureView_Click  '��ʾ�鿴
  
End Sub

Private Sub picShow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 If Button = 2 Then
    PopupMenu MnuPicture
  Else
   '��װ���
  If picShow.ScaleHeight <= picDisplay.Height And picShow.ScaleWidth <= picDisplay.Width Then
     picShow.MouseIcon = picMouseUp.Picture
     Exit Sub
   Else
     picShow.MouseIcon = picMouseDown.Picture
  End If
 End If
  
End Sub

Private Sub picShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If picShow.ScaleHeight <= picDisplay.Height And picShow.ScaleWidth <= picDisplay.Width Then
     Exit Sub
  Else
     MovePicture picShow, X, Y, Button '�ƶ�ͼƬ
  End If

End Sub

Private Sub picShow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  '��װ���
  If picShow.ScaleHeight <= picDisplay.Height And picShow.ScaleWidth <= picDisplay.Width Then
     picShow.MouseIcon = picMouseUp.Picture
   Else
     picShow.MouseIcon = picMouseOver.Picture
  End If
  
End Sub

Private Sub picShow_Resize()

  If picShow.Height > picDisplay.Height Then
     picShow.Top = 0
     picShow.MouseIcon = picMouseOver.Picture
    Else
     picShow.Top = (picDisplay.Height - picShow.Height) / 2
  End If
  If picShow.Width > picDisplay.Width Then
     picShow.Left = 0
     picShow.MouseIcon = picMouseOver.Picture
    Else
     picShow.Left = (picDisplay.Width - picShow.Width) / 2
  End If
  '��װ���
  If picShow.ScaleHeight <= picDisplay.Height And picShow.ScaleWidth <= picDisplay.Width Then
     picShow.MouseIcon = picMouseUp.Picture
  End If
  
End Sub

Private Sub ShellFolderViewOC1_SelectionChanged()

End Sub

Private Sub ShellFolderViewX_SelectionChanged()

End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    On Error Resume Next
    Select Case Button.Key
        Case "��"
          MnuFileOpen_Click
        Case "����"
          AudioDisplay.Run   '����
        Case "����"
           Clipboard.Clear '���������
        Case "����"
            mnuEditCopy_Click
        Case "ɾ��"
            mnuFileDelete_Click
        Case "����"
            MnuFileAttribute_Click
        Case "��ӡ"
            MnuPrintPicture_Click
        Case "��ͼ��"
            mnuView_Click (0)
        Case "Сͼ��"
            mnuView_Click (1)
        Case "�б�"
            mnuView_Click (2)
        Case "��ϸ����"
            mnuView_Click (3)
        Case "����"
            MnuArrangSortAuto_Click
        Case "����"
            MnuArrangSortAutoZ_Click
        Case "������"
            mnuViewToolbar_Click
        Case "״̬��"
            mnuViewStatusBar_Click
        Case "Ԥ����"
            MnuViewPreview_Click
        Case "����"
            
        
    End Select
    
End Sub

Private Sub mnuHelpAbout_Click()
    
    MsgBox "�汾 " & App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    
    Dim nRet As Integer

    '����������û�а����ļ�����ʾ��Ϣ���û�
    '�����ڡ��������ԡ��Ի�����ΪӦ�ó������ð����ļ�
    If Len(App.HelpFile) = 0 Then
        MsgBox "�޷���ʾ����Ŀ¼���ù���û��������İ�����", vbInformation, Me.Caption
    Else

    On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()

    Dim nRet As Integer

    '����������û�а����ļ�����ʾ��Ϣ���û�
    '�����ڡ��������ԡ��Ի�����ΪӦ�ó������ð����ļ�
    If Len(App.HelpFile) = 0 Then
        MsgBox "�޷���ʾ����Ŀ¼���ù���û��������İ�����", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuViewRefresh_Click()

   If picLoad = True Then
      mnuEditCopy.Enabled = False '���ư�ť��Ч
      MnuLookFor.Enabled = False '�鿴�˵���Ч
      MnuPrintPicture.Enabled = False
      '����ͼƬ�˵�
      MnuPicture.Visible = False
      picShow.Visible = False
      tbToolBar.Buttons(Copy_Number).Enabled = False
      tbToolBar.Buttons(Printer_Number).Enabled = False
      sbStatusBar.Panels(3).Text = "δע���: V1.0"
      picLoad = False
   End If
   
   If fPath$ > "" And frmMain.Visible Then
      fPath = ValidateDir(fPath)
      vbGetFileList
   End If
  
   '�˵�Ϊ��Ч
   If lvListView.ListItems.Count > 0 Then
      lvListView.SetFocus  '�б��ý���
    Else
      MenuEnabled (0)
   End If
   
End Sub

Private Sub mnuViewStatusBar_Click()
    
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    
    '״̬���ı�
    If mnuViewStatusBar.Checked = True Then
       tbToolBar.Buttons(Display_Number + 1).Value = tbrPressed
       SaveSetting App.Title, "Settings", "DisplayStatusbar", 1
    Else
       tbToolBar.Buttons(Display_Number + 1).Value = tbrUnpressed
       SaveSetting App.Title, "Settings", "DisplayStatusbar", 0
    End If
    
    SizeControls tvTreeView.Width
    SizeControlsH imgSplitter2.Top
    
End Sub

Private Sub mnuViewToolbar_Click()
    
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
    
    '�������ı�
    If mnuViewToolbar.Checked = True Then
       tbToolBar.Buttons(Display_Number).Value = tbrPressed
       SaveSetting App.Title, "Settings", "DisplayToolbar", 1
    Else
       tbToolBar.Buttons(Display_Number).Value = tbrUnpressed
       SaveSetting App.Title, "Settings", "DisplayToolbar", 0
    End If
    
    SizeControls tvTreeView.Width
    SizeControlsH imgSplitter2.Top
    
End Sub

Private Sub mnuEditCopy_Click()
   
   If picLoad = False Then
     Dim picFile As String
         picFile = fPath$ & lvListView.SelectedItem.Text
     Screen.MousePointer = vbHourglass
     picBuffer.Picture = LoadPicture(picFile)
     Screen.MousePointer = vbDefault
     picLoad = True     '�Ѿ���װ���
   End If
   
   '����ͼƬ��������
   Screen.MousePointer = vbHourglass
   Clipboard.Clear
   Clipboard.SetData picBuffer.Picture, vbCFBitmap
   
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub mnuFileClose_Click()
    
    'ж�ش���
    Unload Me

End Sub

Private Sub mnuFileDelete_Click()
   
   'ɾ���ļ�����
   Dim SHop As SHFILEOPSTRUCT
   Dim strFile As String
   
    strFile = ValidateDir(fPath$) & lvListView.SelectedItem.Text
   
   With SHop
      .wFunc = FO_DELETE
      .pFrom = strFile
      '���Shift������
       If UndoK = True Then
       .fFlags = FOF_NOCONFIRMATION
        Else
       .fFlags = FOF_ALLOWUNDO
       End If
   End With
   
     Dim retVal As Long
         retVal = SHFileOperation(SHop)
         
  'ϵͳɾ���������ʱ,������û��ɾ��
  retVal = GetFileAttributes(strFile)

  If retVal = -1 Then 'ɾ�����ʱ
     lvListView.ListItems.Remove lvListView.SelectedItem.Index
     RefreshDesk
   Else
     Exit Sub 'û��ɾ��ʱ
  End If
   
End Sub

Private Sub MnuFileOpen_Click()
    
 '�򿪵��ļ��Ĵ���
  lvListView_DblClick
    
End Sub

Private Sub tvTreeView_SelectionChange(strPath As String, strDisplayName As String)
   
   If strPath = "" Then Exit Sub  '·��Ϊ��ʱ�˳�
   
   DisplayPath.Text = strPath
       
   RefreshDesk   'ˢ������
   
   fPath$ = strPath
   If fPath$ > "" And frmMain.Visible Then
      fPath = ValidateDir(fPath)
      vbGetFileList
   End If
         
   'ȷ����û�ж���
    If lvListView.ListItems.Count > 0 Then
       MenuEnabled (True)
      Else
       MenuEnabled (False)
    End If
    
End Sub

Private Sub SizeControlsH(X As Single)

    On Error Resume Next

    If picDisplay.Visible = False Then
       tvTreeView.Height = lvListView.Height + picTitles.Height
       imgSplitter2.Visible = False
       Exit Sub
     Else
       imgSplitter2.Visible = True
    End If
    
    '���� Width ����
    If X < 2500 Then X = 2500
    If X > (Me.Height - 2500) Then X = Me.Height - 2500

    imgSplitter2.Width = tvTreeView.Width
    
    Dim Nl As Long
    If tbToolBar.Visible Then
        Nl = tbToolBar.Height
        tvTreeView.Height = X - tbToolBar.Height
        imgSplitter2.Top = X
    Else
        Nl = 0
        tvTreeView.Height = X
        imgSplitter2.Top = X
    End If

    If sbStatusBar.Visible Then
       Nl = Nl + sbStatusBar.Height
    Else
       Nl = Nl
    End If
     
     picDisplay.Height = Me.Height - (tvTreeView.Height + Nl + 810)
     picDisplay.Top = imgSplitter2.Top + 120
  
End Sub

Private Sub ShowFileProperties(ByVal aFile As String)

  '��ʾ�ļ�����
  Dim sei As SHELLEXECUTEINFO
      sei.hWnd = frmMain.hWnd
      sei.lpVerb = "properties"
      sei.lpFile = aFile
      sei.fMask = SEE_MASK_INVOKEIDLIST
      sei.cbSize = Len(sei)
 
  ShellExecuteEx sei

End Sub

Private Sub IniData()
  
   '��ʼ���ϴ�������������
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    imgSplitter.Left = GetSetting(App.Title, "Settings", "Position", 2000)   '��ֱ��
    imgSplitter2.Top = GetSetting(App.Title, "Settings", "HPosition", 1500)  'ˮƽ��
    OldShowSize = Val(GetSetting(App.Title, "Settings", "ShowSize", 0))
    MnuShowSize(OldShowSize).Checked = True '��ʾ��С
    
    sbStatusBar.Panels(3).Text = "δע���: V1.0"
    
    '�����ϴε��Զ������Ƿ���Ч
    mnuArrangeFileIcon.Checked = Val(GetSetting(App.Title, "Settings", "AutoArrange", 0))
    If mnuArrangeFileIcon.Checked Then
       lvListView.Arrange = lvwAutoTop
     Else
       lvListView.Arrange = lvwNone
    End If
    MnuViewPreview.Checked = Val(GetSetting(App.Title, "Settings", "DisplayPreview", 0))
    If MnuViewPreview.Checked Then
       tbToolBar.Buttons(Display_Number + 2).Value = tbrPressed
       picDisplay.Visible = True
     Else
       picDisplay.Visible = False
       tbToolBar.Buttons(Display_Number + 2).Value = tbrUnpressed
    End If
    mnuViewToolbar.Checked = Val(GetSetting(App.Title, "Settings", "DisplayToolbar", 0))
    If mnuViewToolbar.Checked Then
       tbToolBar.Buttons(Display_Number).Value = tbrPressed
       tbToolBar.Visible = True
     Else
       tbToolBar.Visible = False
       tbToolBar.Buttons(Display_Number).Value = tbrUnpressed
    End If
    mnuViewStatusBar.Checked = Val(GetSetting(App.Title, "Settings", "DisplayStatusbar", 0))
    If mnuViewStatusBar.Checked Then
       tbToolBar.Buttons(Display_Number + 1).Value = tbrPressed
       sbStatusBar.Visible = True
     Else
       sbStatusBar.Visible = False
       tbToolBar.Buttons(Display_Number + 1).Value = tbrUnpressed
    End If
    
    '����ͼƬ��ı���ɫΪϵͳ��ɫ
    picShow.BackColor = GetSysColor(4)
    picDisplay.BackColor = picShow.BackColor
    picBuffer.BackColor = picShow.BackColor
    Me.BackColor = picShow.BackColor
    GifView.Left = 20  '��ʼ��λ��
    MnuFullScreen.Checked = False  'ȫ��û��ѡ��
    If AudioDisplay.Visible Then  '���������Чʱ
       AudioDisplay.Left = (picDisplay.Width - AudioDisplay.Width) / 2
       AudioDisplay.Top = (picDisplay.Height - AudioDisplay.Height) / 2
    End If
    
End Sub

Private Sub MenuEnabled(LB As Boolean)

 If LB = False Then
    mnuFileDelete.Enabled = False 'ɾ���˵�
    tbToolBar.Buttons(Copy_Number + 1).Enabled = False
    mnuFileRename.Enabled = False '�������˵�
    mnuEditCopy.Enabled = False   '�����˵�
    MnuLookFor.Enabled = False    '�鿴�˵�
    MnuPrintPicture.Enabled = False '��ӡ�˵�
    tbToolBar.Buttons(Copy_Number).Enabled = False
    tbToolBar.Buttons(Printer_Number).Enabled = False
    mnuEditCopyTo.Enabled = False '�������˵�
    MnuFileOpen.Enabled = False   '�򿪲˵�
    tbToolBar.Buttons(Printer_Number - 2).Enabled = False
    mnuEditMove.Enabled = False   '�ƶ��˵�
    MnuFileAttribute.Enabled = False '����
    tbToolBar.Buttons(2).Enabled = False 'ˢ��
    mnuViewRefresh.Enabled = False
    MnuVideo.Visible = True        'Ӱ�Ӳ˵�
 Else
    mnuFileDelete.Enabled = True  'ɾ���˵�
    tbToolBar.Buttons(Copy_Number + 1).Enabled = True
    mnuFileRename.Enabled = True  '������
    mnuEditCopyTo.Enabled = True  '������
    MnuFileOpen.Enabled = True    '��
    tbToolBar.Buttons(Printer_Number - 2).Enabled = True
    mnuEditMove.Enabled = True    '�ƶ�
    MnuFileAttribute.Enabled = True '����
    tbToolBar.Buttons(2).Enabled = True '�򿪰�ť
    mnuViewRefresh.Enabled = True
 End If
      
End Sub

Private Sub PictureProccess(picFile As String)
 
      IsPicture (True)
     '����ͼ���W��H,����װͼƬ
      sbStatusBar.Panels(3).Text = "��=" & GetImageSize(picFile).Width & " ��=" & _
                                   GetImageSize(picFile).Height
      mnuEditCopy.Enabled = True '���ư�ť��Ч
      MnuLookFor.Enabled = True  '�鿴�˵���Ч
      MnuPrintPicture.Enabled = True
      tbToolBar.Buttons(Copy_Number).Enabled = True
      tbToolBar.Buttons(Printer_Number).Enabled = True
      '��װͼƬ
      On Error GoTo Nopic
      Screen.MousePointer = vbArrowHourglass
      '�Ƿ�װͼƬ
      If picDisplay.Visible Then
         picBuffer.Picture = LoadPicture(picFile)
         picLoad = True  '�Ѿ���װͼƬ
         ShowPreview picBuffer, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picShow, picDisplay.ScaleWidth, picDisplay.ScaleHeight, picDisplay.Visible, OldShowSize
       Else
         picLoad = False  'û�а�װ
      End If
      Screen.MousePointer = vbDefault
      
      Exit Sub
     '���ͼƬ
Nopic:
      picBuffer.Picture = LoadPicture()
      picShow.Cls
      picShow.Picture = LoadPicture()
      picLoad = True
      Screen.MousePointer = vbDefault
      mnuEditCopy.Enabled = False '���ư�ť��Ч
      MnuLookFor.Enabled = False  '�鿴�˵���Ч
      MnuPrintPicture.Enabled = False
     '����ͼƬ�˵�
      IsPicture (False)
      tbToolBar.Buttons(Copy_Number).Enabled = False
      tbToolBar.Buttons(Printer_Number).Enabled = False
      sbStatusBar.Panels(3).Text = "δע���: V1.0"
 
End Sub

Private Sub GifProccess(picFile As String)

  If picDisplay.Visible Then  'Ԥ����Чʱ
     IsGif (True)
    '����ͼ���W��H,����װͼƬ
     sbStatusBar.Panels(3).Text = "��=" & GetImageSize(picFile).Width & " ��=" & _
                          GetImageSize(picFile).Height
     GifView.Navigate picFile  '��ʾ�����ļ�
     GifView.Height = picDisplay.ScaleHeight - 40
     GifView.Width = picDisplay.ScaleWidth - 40
  End If
  
End Sub

Private Sub RefreshDesk()

  If picShow.Visible Then
      IsPicture (False)
      mnuEditCopy.Enabled = False '���ư�ť��Ч
      MnuLookFor.Enabled = False  '�鿴�˵���Ч
      MnuPrintPicture.Enabled = False
     '����ͼƬ�˵�
      tbToolBar.Buttons(Copy_Number).Enabled = False
      tbToolBar.Buttons(Printer_Number).Enabled = False
  End If
  
  If GifView.Visible Then  'GIF���ʱ
     IsGif (False)
  End If
  
  If AudioDisplay.Visible Then  '������Ƶ
       tbToolBar.Buttons(4).Enabled = False
     AudioDisplay.FileName = App.Path + "\camera.wav"
     AudioDisplay.Run
     IsAudio (False)
     frmMain.MnuVideo.Visible = False
  End If
  
   sbStatusBar.Panels(3).Text = "δע���: V1.0"
   sbStatusBar.Panels(4).Text = ""
   sbStatusBar.Panels(5).Text = ""
   sbStatusBar.Panels(6).Text = ""
   sbStatusBar.Panels(7).Text = ""
   sbStatusBar.Panels(8).Text = ""
   
End Sub

Private Sub txtProccess(picFile As String)
 
  On Error GoTo ErrNo
  If picDisplay.Visible Then  'Ԥ����Чʱ
     IsGif (True)
     sbStatusBar.Panels(3).Text = "δע���: V1.0"
     GifView.Navigate picFile '��ʾ�ı��ļ�
     GifView.Height = picDisplay.ScaleHeight - 40
     GifView.Width = picDisplay.ScaleWidth - 40
  End If
  
  Exit Sub
ErrNo:
  GifView.Visible = False
  
End Sub

Private Sub AudioProccess(picFile As String)
 
  On Error GoTo ErrNo
  If picDisplay.Visible Then  'Ԥ����Чʱ
     IsAudio (True)
     sbStatusBar.Panels(3).Text = "δע���: V1.0"
     'ȷ���Ƿ��Զ�����
     If Val(GetSetting(App.Title, "Settings", "AutoPlay", 1)) = 1 Then
        AudioDisplay.AutoStart = True
      Else
        AudioDisplay.AutoStart = False
     End If
     AudioDisplay.FileName = picFile  '��λ�ļ�
  End If
  
  Exit Sub
ErrNo:
  AudioDisplay.Visible = False
  tbToolBar.Buttons(4).Enabled = False
  
End Sub

Private Sub VideoProccess(picFile As String)
 
  On Error GoTo ErrNo
  If picDisplay.Visible Then  'Ԥ����Чʱ
     IsAudio (True)
     sbStatusBar.Panels(3).Text = "δע���: V1.0"
     'ȷ���Ƿ��Զ�����
     If Val(GetSetting(App.Title, "Settings", "AutoPlay", 1)) = 1 Then
        AudioDisplay.AutoStart = True
      Else
        AudioDisplay.AutoStart = False
     End If
     AudioDisplay.FileName = picFile  '��λ�ļ�
     MnuVideo.Visible = True
  End If
  
  Exit Sub
ErrNo:
  AudioDisplay.Visible = False
  MnuVideo.Visible = False
  
End Sub


