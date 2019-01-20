VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "测试UI界面"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9075
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

Private m_MouseSurface As CSurface
Private m_UIManger As CUIManager
Private m_UIForm As CUIForm
Private m_UIForm2 As CUIForm
Private m_testCmd As CUICommand
Private m_testCmd2 As CUICommand

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
        .SetRefreshSpeed 60
        .SetScreen Width / Screen.TwipsPerPixelX, _
                   Height / Screen.TwipsPerPixelX, _
                    hWnd
    End With
    'Clear All Resoures
    End
End Sub

Private Sub m_VBGameEngine_GameExit()
    '游戏的资源释放销毁
    Set m_MouseSurface = Nothing
    Set m_UIManger = Nothing
    Set m_UIForm = Nothing
    Set m_UIForm2 = Nothing
    Set m_testCmd = Nothing
    Set m_testCmd2 = Nothing
    Set m_VBGameEngine = Nothing
End Sub

Private Sub m_VBGameEngine_GameInit(GameInit As Boolean)
    '引擎初始化代码

    Set m_MouseSurface = New CSurface
    m_MouseSurface.LoadJPG GetResPath + "cursor.bmp"
    
    Set m_UIManger = New CUIManager
    Set m_UIForm = New CUIForm
    m_UIForm.Create 20, 20, 485, 380
    m_UIManger.RegWindow m_UIForm
    
    Set m_UIForm2 = New CUIForm
    m_UIForm2.Create 40, 40, 485, 380
    m_UIManger.RegWindow m_UIForm2
    
    Set m_testCmd = New CUICommand
    Set m_testCmd.UIControl.Parent = m_UIForm
    m_testCmd.Caption = " "
    m_testCmd.UIControl.Create 20, 20, 60, 20
    
    Set m_testCmd2 = New CUICommand
    Set m_testCmd2.UIControl.Parent = m_UIForm
    m_testCmd2.Caption = " "
    m_testCmd2.UIControl.Create 20, 80, 60, 20
End Sub

Private Sub m_VBGameEngine_GameRefresh()
    '刷新屏幕
    g_MainSurface.Clear
    m_UIManger.Render
    g_MainSurface.Blt m_MouseSurface, g_Inputs.MouseX, g_Inputs.MouseY
    
    If g_Inputs.KeyDown(1) Then m_VBGameEngine.ExitGame = True
End Sub

Private Sub m_VBGameEngine_InputMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)
    m_UIManger.SendInputMsg MsgType, Value, x, y
    g_MainSurface.TextOut 0, 20, Hex(MsgType)
End Sub
