VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   0  'None
   Caption         =   "测试DInput"
   ClientHeight    =   5490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_VBGameEngine As CGameEngine
Attribute m_VBGameEngine.VB_VarHelpID = -1
Private m_Video As CVedioAudio

Private m_x As Integer
Private m_y As Integer

Private Sub Form_Unload(Cancel As Integer)
    '注意销毁的代码
    Cancel = 1
    m_VBGameEngine.ExitGame = True
End Sub

Private Sub Form_Load()
    Me.Show
    Set m_VBGameEngine = New CGameEngine
    With m_VBGameEngine
        .SetMode 0
        .MouseVisiable = False
        .SetRefreshSpeed 85
        .SetScreen 800, 600, hWnd, False
        '.SetScreen Width / Screen.TwipsPerPixelX, _
                   Height / Screen.TwipsPerPixelX, _
                    hWnd
    End With
    'Clear All Resoures
    End
End Sub

Private Sub m_VBGameEngine_GameExit()
    '游戏的资源释放销毁
    Set m_Video = Nothing
    Set m_VBGameEngine = Nothing
End Sub

Private Sub m_VBGameEngine_GameInit(GameInit As Boolean)
    '引擎初始化代码
    m_x = 115
    m_y = 100
    Set m_Video = New CVedioAudio
    Dim VideoRect As RECT
    VideoRect.Left = 0
    VideoRect.Top = 50
    VideoRect.Right = 800
    VideoRect.Bottom = 550
    m_Video.OpenMedia GetResPath + "IE.avi"
    
    m_Video.SetVedioRect hWnd, VideoRect
    m_Video.PlayMedia
    Set m_Video = Nothing
End Sub

Private Sub m_VBGameEngine_GameRefresh()
    '刷新屏幕
    g_MainSurface.Clear
    g_MainSurface.TextOut 0, 15, "Press ESC To Exit!", vbGreen
    g_MainSurface.TextOut 0, 30, "Press W A S D To Move Cursor！", vbGreen

    If g_Inputs.KeyDown(1) Then m_VBGameEngine.ExitGame = True
End Sub
