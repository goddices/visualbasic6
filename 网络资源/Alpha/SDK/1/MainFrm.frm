VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SDK1"
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
'**��    �ڣ�2005-07-05 18:59:25
'**�� �� �ˣ�
'**��    �ڣ�
'**��    �������濪��ʹ��ʵ����
'**��    ����V1.0.0
'*************************************************************************
Option Explicit

Private WithEvents m_VBGameEngine As CGameEngine
Attribute m_VBGameEngine.VB_VarHelpID = -1

Private Sub Form_Unload(Cancel As Integer)
    'ע�����ٵĴ���
    Cancel = 1
    m_VBGameEngine.ExitGame = True
End Sub

Private Sub Form_Load()
    Me.Show
    Set m_VBGameEngine = New CGameEngine
    With m_VBGameEngine
        .SetMode 1
        .SetRefreshSpeed 60
        .MouseVisiable = True
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
End Sub

Private Sub m_VBGameEngine_GameInit(GameInit As Boolean)
    '�����ʼ������
    
End Sub

Private Sub m_VBGameEngine_GameRefresh()
    'ˢ����Ļ
    g_MainSurface.Clear
    g_MainSurface.TextOut 0, 20, "��ã�����"
End Sub
