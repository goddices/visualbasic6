VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "测试DSound"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_VBGameEngine As CGameEngine
Attribute m_VBGameEngine.VB_VarHelpID = -1

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
        .SetScreen Width / Screen.TwipsPerPixelX, _
                   Height / Screen.TwipsPerPixelX, _
                    hWnd
    End With
    'Clear All Resoures
    End
End Sub

Private Sub m_VBGameEngine_GameExit()
    '游戏的资源释放销毁
    Set m_VBGameEngine = Nothing
End Sub

Private Sub m_VBGameEngine_GameInit(GameInit As Boolean)
    '引擎初始化代码
    g_Sounds.LoadWav GetResPath + "0.wav", "Sound0"
    g_Sounds.LoadWav GetResPath + "1.wav", "Sound1"
End Sub

Private Sub m_VBGameEngine_GameRefresh()
    '刷新屏幕
    g_MainSurface.Clear
    g_MainSurface.TextOut 100, 100, "Click Mouse To Play Sound！"
End Sub

Private Sub m_VBGameEngine_InputMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)
    If MsgType = MSG_MouseClick Then
        If Value = 1 Then g_Sounds.Play "Sound0"
        If Value = 2 Then g_Sounds.Play "Sound1"
    End If
End Sub
