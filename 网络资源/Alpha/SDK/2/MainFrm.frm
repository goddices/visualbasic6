VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SDK2"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************************
'**ģ �� ����MainFrm
'**˵    ����Boywhp ��Ȩ����2005 - 2006(C)
'**�� �� �ˣ�����ƽ
'**��    �ڣ�2005-07-05 19:00:32
'**�� �� �ˣ�
'**��    �ڣ�
'**��    �������濪��ʵ����
'**��    ����V1.0.0
'*************************************************************************

Option Explicit

Private WithEvents m_VBGameEngine As CGameEngine
Attribute m_VBGameEngine.VB_VarHelpID = -1

Private m_Sprite As CSurface
Private m_CursorX As Integer
Private m_CursorY As Integer

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_CursorX = x
    m_CursorY = y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'ע�����ٵĴ���
    Cancel = 1
    m_VBGameEngine.ExitGame = True
End Sub

Private Sub Form_Load()
    Debug.Print GetResPath
    Me.Show
    Me.ScaleMode = 3
    
    Set m_VBGameEngine = New CGameEngine
    With m_VBGameEngine
        .SetMode 1
        .MouseVisiable = True
        .SetRefreshSpeed 60
        .SetScreen Me.Width / Screen.TwipsPerPixelX, _
                    Me.Height / Screen.TwipsPerPixelY, _
                    hWnd
    End With
    'Clear All Resoures
    End
End Sub

Private Sub m_VBGameEngine_GameExit()
    '��Ϸ����Դ�ͷ�����
    Set m_VBGameEngine = Nothing
    Set m_Sprite = Nothing
End Sub

Private Sub m_VBGameEngine_GameInit(GameInit As Boolean)
    '�����ʼ������
    Set m_Sprite = New CSurface
    m_Sprite.LoadBMP GetResPath + "cursor.bmp", 28, 28
End Sub

Private Sub m_VBGameEngine_GameRefresh()
    'ˢ����Ļ
    g_MainSurface.Clear
    g_MainSurface.Blt m_Sprite, m_CursorX, m_CursorY
End Sub
