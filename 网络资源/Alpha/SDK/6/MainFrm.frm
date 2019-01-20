VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "测试DInput"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7275
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

Private m_x As Integer
Private m_y As Integer
Private m_z As Integer
Private m_Sprite As CSurface

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
        .MouseVisiable = False
        .SetRefreshSpeed 85
        .SetScreen Width / Screen.TwipsPerPixelX, _
                   Height / Screen.TwipsPerPixelX, _
                    hWnd
    End With
    'Clear All Resoures
    End
End Sub

Private Sub m_VBGameEngine_GameExit()
    '游戏的资源释放销毁
    Set m_Sprite = Nothing
    Set m_VBGameEngine = Nothing
End Sub

Private Sub m_VBGameEngine_GameInit(GameInit As Boolean)
    '引擎初始化代码
    m_x = 115
    m_y = 100
    
    Set m_Sprite = New CSurface
    m_Sprite.LoadJPG GetResPath + "mm.bmp"
End Sub

Private Sub m_VBGameEngine_GameRefresh()
    '刷新屏幕
    g_MainSurface.ColorFill 100, 100, 100

    Dim i As Integer
    For i = 0 To 5
        g_MainSurface.Blt m_Sprite, i * 80, 0, i + 1
    Next
    g_MainSurface.Blt m_Sprite, 0, 80
    g_MainSurface.Blt m_Sprite, 0, m_z / 5, 9
    
    g_MainSurface.TextOut 0, 15, "Press ESC To Exit!", vbGreen
    g_MainSurface.TextOut 0, 30, "Press W A S D To Move Cursor！", vbGreen
    g_MainSurface.TextOut m_x, m_y, "+"
    If g_Inputs.KeyDown(17) Then m_y = m_y - 1
    If g_Inputs.KeyDown(31) Then m_y = m_y + 1
    If g_Inputs.KeyDown(30) Then m_x = m_x - 1
    If g_Inputs.KeyDown(32) Then m_x = m_x + 1
    
    If g_Inputs.KeyDown(1) Then m_VBGameEngine.ExitGame = True
    
End Sub

Private Sub m_VBGameEngine_InputMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)
    If MsgType = MSG_MouseMove Then
        m_x = x
        m_y = y
    End If
    If MsgType = MSG_MouseScroll Then
        m_z = m_z + Value
    End If
    If MsgType = MSG_MouseDblClick Then
        m_VBGameEngine.ExitGame = True
    End If
End Sub
