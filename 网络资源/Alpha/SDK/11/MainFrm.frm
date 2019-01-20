VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   0  'None
   Caption         =   "测试DInput"
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5700
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

Private m_x As Long
Private m_y As Long

Private m_BS As CSurface
Private m_LS As CSurface
Private m_Custor As CSurface

Private Sub Form_Unload(Cancel As Integer)
    '注意销毁的代码
    Cancel = 1
    m_VBGameEngine.ExitGame = True
End Sub

Private Sub Form_Load()
    Set m_VBGameEngine = New CGameEngine
    With m_VBGameEngine
        .SetMode 1
        .MouseVisiable = False
        .SetRefreshSpeed 60
        .SetScreen 640, _
                   480, _
                    hWnd, False
    End With
    'Clear All Resoures
    End
End Sub

Private Sub m_VBGameEngine_GameExit()
    '游戏的资源释放销毁
    Set m_BS = Nothing
    Set m_LS = Nothing
    Set m_Custor = Nothing
    Set m_VBGameEngine = Nothing
End Sub

Private Sub m_VBGameEngine_GameInit(GameInit As Boolean)
    '引擎初始化代码
    m_x = 115
    m_y = 100
    
    Set m_BS = New CSurface
    m_BS.LoadBMP GetResPath + "mm.bmp", 576, 476
    Set m_LS = New CSurface
    m_LS.LoadJPG GetResPath + "light.bmp"
    
    Set m_Custor = New CSurface
    m_Custor.LoadJPG GetResPath + "cursor.bmp"
    
    g_Screen.CreateLightTable
    g_Screen.EnableZBuffer
End Sub

Private Sub m_VBGameEngine_GameRefresh()
    '刷新屏幕
    g_MainSurface.Clear
    
    g_Screen.ClearZBuffer
    g_Screen.BltWithZBuffer m_BS, 0, 0, 101, 0
    g_Screen.BltWithZBuffer m_BS, 100, 0, 100, 0
    g_Screen.BltWithZBuffer m_BS, 200, 0, 1000, 0
    g_Screen.BltWithZBuffer m_BS, 300, 0, 1001, 0
    g_Screen.BltWithZBuffer m_BS, 400, 0, 1002, 0
    
    If g_Screen.GetZBuffer(m_x, m_y) = 101 Then
        g_MainSurface.BltWithEdgeline m_BS, 0, 0, &H1F
        g_MainSurface.TextOut 0, 40, "发现当前目标：—)", vbGreen
   End If
   
   '进行光泽处理
    g_Screen.SetAmbientLight 80 + 20 * Sin(Timer)
    g_Screen.BltToLightTable m_LS, -64, -64
    g_Screen.BltToLightTable m_LS, m_x - 64, m_y - 64
    
    g_MainSurface.TextOut 0, 20, "Z_Buffer:" + Str(g_Screen.GetZBuffer(m_x, m_y)), vbGreen
    g_MainSurface.FastBltEx m_Custor, m_x, m_y
    If g_Inputs.KeyDown(1) Then m_VBGameEngine.ExitGame = True
End Sub

Private Sub m_VBGameEngine_InputMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)
    If MsgType = MSG_MouseMove Then
        m_x = x
        m_y = y
    End If
End Sub
