VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SDK10"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture1 
      Height          =   4605
      Left            =   90
      MousePointer    =   2  'Cross
      ScaleHeight     =   303
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   387
      TabIndex        =   0
      Top             =   90
      Width           =   5865
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_VBGameEngine As CGameEngine
Attribute m_VBGameEngine.VB_VarHelpID = -1

Private m_Sprite As CSurface
Private m_RoteBuffer As CSurface
Private m_CursorX As Integer
Private m_CursorY As Integer

Private Sub Form_Unload(Cancel As Integer)
    '注意销毁的代码
    Cancel = 1
    m_VBGameEngine.ExitGame = True
End Sub

Private Sub Form_Load()
    Me.Show
    Set m_VBGameEngine = New CGameEngine
    With m_VBGameEngine
        .SetMode 1
        .MouseVisiable = True
        .SetRefreshSpeed 60
        .SetScreen Picture1.Width / Screen.TwipsPerPixelX, _
                    Picture1.Height / Screen.TwipsPerPixelX, _
                    hWnd, , _
                    Picture1.hWnd
                    
    End With
    'Clear All Resoures
    End
End Sub

Private Sub m_VBGameEngine_GameExit()
    '游戏的资源释放销毁
    Set m_Sprite = Nothing
    Set m_RoteBuffer = Nothing
    Set m_VBGameEngine = Nothing
End Sub

Private Sub m_VBGameEngine_GameInit(GameInit As Boolean)
    '引擎初始化代码
    Set m_Sprite = New CSurface
    m_Sprite.LoadBMP GetResPath + "cursor2.bmp", 200, 125
    
    Set m_RoteBuffer = New CSurface
    m_RoteBuffer.Create 150, 150
End Sub

Private Sub m_VBGameEngine_GameRefresh()
    '刷新屏幕
    g_MainSurface.Clear
    
    m_RoteBuffer.ColorFill 255, 0, 255
    m_RoteBuffer.RotateRect m_Sprite, 0, 0, Timer * 2, 1
    
    g_MainSurface.AddColorEx m_RoteBuffer, m_CursorX - 63, m_CursorY - 63, RGBtoDDColor((Sin(Timer) + 1) / 2 * 100, 0, 0)
    g_MainSurface.FastAdditive m_Sprite, m_CursorX, m_CursorY
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_CursorX = x
    m_CursorY = y
End Sub
