VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SDK9"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
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

Private m_x As Long
Private m_y As Long
Private m_z As Long
Private m_Lei As CSurface
Private m_Yun As CSurface

Dim Yun_i   As Long      '背景云的位置X
Dim Yun1_X()  As Long    '云1的位置X
Dim Yun1_Y()  As Long    '云1的位置Y
Dim Yun2_X()  As Long    '云2的位置Y
Dim Yun2_Y()  As Long    '云2的位置Y
Dim Lei_X() As Long      '火箭的位置X
Dim Lei_Y() As Long      '火箭的位置Y
Dim Lei_A() As Long      '火箭的动作

Dim i   As Long          '临时变量
Dim timeCount   As Long  '计时器
Dim timeCount1  As Long  '计时器1
Dim Speed1  As Long      '速度

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
        .SetRefreshSpeed 60
        .SetScreen 640, 480, Me.hWnd, False
        '.SetScreen Width / Screen.TwipsPerPixelX, _
                   Height / Screen.TwipsPerPixelX, _
                    hWnd, True
    End With
    'Clear All Resoures
    End
End Sub

Private Sub m_VBGameEngine_GameExit()
    '游戏的资源释放销毁
    Set m_Lei = Nothing
    Set m_Yun = Nothing
    Set m_VBGameEngine = Nothing
End Sub

Private Sub m_VBGameEngine_GameInit(GameInit As Boolean)
    '引擎初始化代码
    m_x = 115
    m_y = 100
    
    Set m_Lei = New CSurface
    m_Lei.LoadJPG GetResPath + "lei.bmp"
    Set m_Yun = New CSurface
    m_Yun.LoadJPG GetResPath + "\Yun.bmp"
    
    Yun_i = 0
    Speed1 = 1
    
    ReDim Yun1_X(50)
    ReDim Yun1_Y(50)
    ReDim Yun2_X(50)
    ReDim Yun2_Y(50)
    For i = 1 To 50
        Randomize
        Yun1_X(i) = Int((800) * Rnd + 640)
        Randomize
        Yun2_X(i) = Int((800) * Rnd + 640)
        Randomize
        Yun1_Y(i) = Int((400) * Rnd + 1)
        Randomize
        Yun2_Y(i) = Int((400) * Rnd + 1)
    Next i
    
    ReDim Lei_X(10)
    ReDim Lei_Y(10)
    ReDim Lei_A(10)
    For i = 1 To 10
        Randomize
        Lei_X(i) = Int((-128) * Rnd - 128)
        Randomize
        Lei_Y(i) = Int((460) * Rnd + 1)
        Randomize
        Lei_A(i) = Int((4) * Rnd + 1)
    Next i
    
End Sub

Private Sub m_VBGameEngine_GameRefresh()
    '刷新屏幕
    g_MainSurface.ColorFill 100, 100, 100

    Dim i As Integer
    g_MainSurface.Blt m_Yun, Yun_i, 0, 0, 4
    
    If Speed1 > 0 Then
        If Yun_i <= -1024 Then Yun_i = 0
        g_MainSurface.Blt m_Yun, Yun_i + 1024, 0, 0, 4
        g_MainSurface.Blt m_Yun, Yun_i - 1024, 0, 0, 4
    ElseIf Speed1 < 1 Then
        If Yun_i >= 1024 Then Yun_i = 0
        g_MainSurface.Blt m_Yun, Yun_i + 1024, 0, 0, 4
        g_MainSurface.Blt m_Yun, Yun_i - 1024, 0, 0, 4
    End If
    Yun_i = Yun_i - Speed1
        
    For i = 1 To 10
        g_MainSurface.Additive m_Yun, Yun1_X(i), Yun1_Y(i), 100, 1
        g_MainSurface.Additive m_Yun, Yun2_X(i), Yun2_Y(i), 80, 2
    
        Yun1_X(i) = Yun1_X(i) - Speed1 * 2
        If Speed1 > 0 Then
            If Yun1_X(i) < -200 Then
                Randomize
                Yun1_Y(i) = Int((400) * Rnd + 1)
                Yun1_X(i) = 640
            End If
        Else
            If Yun1_X(i) > 640 Then
                Randomize
                Yun1_Y(i) = Int((400) * Rnd + 1)
                Yun1_X(i) = -200
            End If
        End If
        
        Yun2_X(i) = Yun2_X(i) - Speed1 * 3
        If Speed1 > 0 Then
            If Yun2_X(i) < -50 Then
                Randomize
                Yun2_Y(i) = Int((400) * Rnd + 1)
                Yun2_X(i) = 640
            End If
        Else
            If Yun2_X(i) > 640 Then
                Randomize
                Yun2_Y(i) = Int((400) * Rnd + 1)
                Yun2_X(i) = -50
            End If
        End If
    Next i
    
    For i = 1 To 10
        g_MainSurface.Additive m_Lei, Lei_X(i), Lei_Y(i), 150, Lei_A(i)
        timeCount = timeCount + 1
        Lei_X(i) = Lei_X(i) + Int((10) * Rnd + 8) - Speed1 / 2
        If Lei_X(i) > 640 Then
            Lei_X(i) = -128
            Randomize
            Lei_Y(i) = Int((460) * Rnd + 1)
        End If
        If timeCount > 10 Then
            Lei_A(i) = Lei_A(i) - 1
            If Lei_A(i) < 0 Then

                Randomize
                Lei_A(i) = Int((4) * Rnd + 1)
                Lei_A(i) = 2
            End If
            timeCount = 0
        End If
    Next i
    
    
    g_MainSurface.TextOut 0, 15, "Press ESC To Exit!", vbGreen
    g_MainSurface.TextOut 0, 30, "Press MouseLeft MouseRight or MouseScroll To Speed Up or Down！", vbGreen
    g_MainSurface.TextOut 0, 45, "Speed:" & Speed1, vbGreen
    g_MainSurface.TextOut m_x, m_y, "+"
    
    If g_Inputs.KeyDown(1) Then m_VBGameEngine.ExitGame = True
    
End Sub

Private Sub m_VBGameEngine_InputMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)
    If MsgType = MSG_MouseMove Then
        m_x = x
        m_y = y
    End If
    If MsgType = MSG_MouseScroll Then
        Speed1 = Speed1 + Value / 100
        If Speed1 > 20 Then Speed1 = 20
        If Speed1 < -20 Then Speed1 = -20
    End If
    'If MsgType = MSG_MouseDblClick Then
    '    m_VBGameEngine.ExitGame = True
    'End If
    If MsgType = MSG_MouseClick And Value = 1 Then
        Speed1 = Speed1 + 1
        If Speed1 > 20 Then Speed1 = 20
    End If
    If MsgType = MSG_MouseClick And Value = 2 Then
        Speed1 = Speed1 - 1
        If Speed1 < -20 Then Speed1 = -20
    End If
End Sub
