VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGameEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===========================Class GameEngine============================
'   SetScreen(Width,Height,Hwnd,Windowned):
'   设置视窗的Width,Height,Hwnd，和运行模式：全屏或窗体
'   SetRefreshSpeed(RefreshSpeed):
'   设置游戏的刷新率Fps。
'=======================================================================
Private Declare Function GetInputState Lib "user32" () As Long

Private WithEvents m_Input As CInput
Attribute m_Input.VB_VarHelpID = -1
Private m_SoundWav As CSoundWav                 '共享

Private m_ExitGame As Boolean
Private m_ShowSysMouse As Boolean               '是否显示系统鼠标
Private m_AutoGameLoop As Boolean               '是否自动开始主循环
Private m_SmoothMove As Boolean
Private m_SoundUsed As Boolean
Private m_InputsUsed As Boolean
Private m_Font As StdFont

Public Event GameInit(GameInit As Boolean)
Public Event GameLoop()
Public Event GameExit()
Public Event GameRefresh()

Public Event InputMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)

Private Sub Destroy()
    Set m_Font = Nothing
    Set g_Screen = Nothing
    Set m_SoundWav = Nothing
    Set g_MainSurface = Nothing
    Set g_DD = Nothing
    Set g_DX7 = Nothing
End Sub

Private Sub Class_Initialize()
    If Not InitDX7 Then
        'Use Our GDI Engine To TakeUp
        MsgBox "请确认安装了DirectX7.0或以上版本！", , "GameEngine"
        Err.Raise vbObjectError + 513, "VBGameEngine", "InitDX7 Failed!"
        Exit Sub
    End If

    Set g_Screen = New CScreen
    Set m_Input = New CInput
    Set m_SoundWav = New CSoundWav
    Set g_MainSurface = New CSurface
    Set m_Font = New StdFont
    
    g_Mode = 0
    m_ShowSysMouse = False
    m_AutoGameLoop = True
    m_SmoothMove = True   ' False
    
    SetRefreshSpeed 50
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Sub SetMode(mode As Long)
    g_Mode = mode
End Sub

Public Function TickCount() As Long
    TickCount = GetTickCount
End Function

Public Function DrawSurface() As DirectDrawSurface7
    Set DrawSurface = g_Screen.DrawSurface
End Function

Public Property Get Screen() As CScreen
    Set Screen = g_Screen
End Property

Public Property Let ExitGame(ByVal IsExitGame As Boolean)
    m_ExitGame = IsExitGame
End Property

Public Property Let AutoGameLoop(ByVal vNewValue As Boolean)
    m_AutoGameLoop = vNewValue
End Property

Public Property Get MouseVisiable() As Boolean
    MouseVisiable = m_ShowSysMouse
End Property

Public Property Let MouseVisiable(ByVal Visiable As Boolean)
    m_ShowSysMouse = Visiable
End Property

Public Function viewRect() As RECT
    viewRect = g_ViewRect
End Function

Public Function MainSurface() As CSurface
    Set MainSurface = g_MainSurface
End Function

Public Function SetScreen(ByVal Width As Long, ByVal Height As Long, ByVal hWnd As Long, _
    Optional windowed As Boolean = True, Optional view_hwnd As Long, Optional flags As EngineSetupConst = 3) As Boolean
    'On Error Resume Next
    
    g_Hwnd = hWnd
    g_Windowed = windowed
    m_SoundUsed = getbit(flags, 1)
    m_InputsUsed = getbit(flags, 2)
    
    If view_hwnd = 0 Then view_hwnd = g_Hwnd
    '安装Screen
    Width = Width And &HFFFFFFFC
    
    If Not InitScreen(Width, Height, view_hwnd) Then
        '严重错误！必须终止程序！
        Destroy
        Err.Raise vbObjectError + 513, "VBGameEngine", "初始化DirectDraw出错!"
        Exit Function
    End If
    
    '初始化字体
    InitFont
    
    If m_InputsUsed Then
        If Not m_Input.Create(m_ShowSysMouse) Then
            'Err To Input
            g_Err_Description = "游戏初始化输入设备出错!"
            MsgBox g_Err_Description, vbOKOnly, "GameEngine"
            Destroy
            Exit Function
        End If
        Set g_Inputs = m_Input
    End If
     
    If m_SoundUsed Then
        If Not m_SoundWav.Create(g_Hwnd) Then
            g_Err_Description = "初始化声音设备出错！"
            'MsgBox g_Err_Description, vbOKOnly, "GameEngine"
            'Destrory
            'Exit Function
        End If
    End If
    Set g_Sounds = m_SoundWav
    
    SetScreen = True
    
    Dim Game_Init As Boolean
    Game_Init = True
    RaiseEvent GameInit(Game_Init)
    If Not Game_Init Then
        g_Err_Description = "游戏初始化数据出错!"
        MsgBox g_Err_Description, vbOKOnly, "GameEngine"
        SetScreen = False
        Destroy
        Exit Function
    End If
    
    'Start GameLoop
    If m_AutoGameLoop Then StartGame
End Function

Public Property Get Inputs() As CInput
    Set Inputs = m_Input
End Property

Public Sub SetRefreshSpeed(RefreshSpeed As Long)
    '10 -> Min_FPS
    Debug.Assert Not g_Screen Is Nothing
    If RefreshSpeed > 10 Then
        g_Screen.RefreshDelay = 1000 / RefreshSpeed
    End If
End Sub

Private Sub StartGame()
    While Not m_ExitGame
        RaiseEvent GameLoop
        'If GetInputState Then DoEvents
        DoEvents
        
        #If EnableRefreshSpeed Then
            '使用帧速控制，关闭撕裂
            If g_Screen.TimeOut Then
                If m_InputsUsed Then m_Input.Refresh
                
                g_Screen.Render   'Draw Last Buffer
                RaiseEvent GameRefresh
                
                #If RealseVersion = 0 Then
                    g_MainSurface.TextOut 0, 0, "FPS:" + Str(g_Screen.FPS), vbGreen
                #End If
            End If
        #Else
            '全速工作，达到系统刷新率
            If g_DD.GetVerticalBlankStatus <> 0 Then
                If m_InputsUsed Then m_Input.Refresh
                
                g_Screen.Render   'Draw Last Buffer
                RaiseEvent GameRefresh
                '#If RealseVersion = 0 Then
                    '实际帧数
                    g_MainSurface.TextOut 0, 0, "FPS:" + Str(g_Screen.FPS), vbGreen
                '#End If
            End If
        #End If
        
        While Not CheckActived And Not m_ExitGame
            DoEvents
            Sleep 1
        Wend
    Wend
    
    'ExitGame Destroy objects
    Destroy
    RaiseEvent GameExit
End Sub

Public Sub Refresh()
    g_Screen.Render
End Sub

Private Sub m_Input_InputMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)
    RaiseEvent InputMsg(MsgType, Value, x, y)
End Sub

Public Function Sounds() As CSoundWav
    Set Sounds = m_SoundWav
End Function

Private Function InitDX7() As Boolean
    Set g_DX7 = New DirectX7
    InitDX7 = Not g_DX7 Is Nothing
End Function

Private Function InitScreen(Width As Long, Height As Long, view_hwnd As Long) As Boolean
    If g_Mode = 0 Then
        'Default DX
        g_MainSurface.DD_Surface = g_Screen.Create(view_hwnd, Width, Height, , g_Windowed)
    Else
        g_MainSurface.DD_Surface = g_Screen.Create(view_hwnd, Width, Height, , g_Windowed, True)
    End If
    
    If g_MainSurface.DD_Surface Is Nothing Then
        Exit Function
    End If
    
    g_MainSurface.SetClipper 0, 0, Width, Height
    g_MainSurface.ReSize Width, Height
    g_MainSurface.AddFrame g_MainSurface.GetClipper, 0, 0
    InitScreen = True
End Function

Private Function InitFont() As Boolean
    '初始化字体
    'bool createfont( //创建逻辑字体，成功返回true,否则返回false。
    'int nheight, // 所创建字体的字符高度
    'int nwidth, // 字体的字符平均宽度
    'int nescapement, // 字符输出方向与水平向右的方向所成角度，
    '以0.1度为单位
    'int norientation, // 字符与基线的角度，以0.1度为单位
    'int nweight, // 字符颜色的深浅度
    'byte bitalic, // 斜体属性标志(0:正常字体，非0:斜体)
    'byte bunderline, //下划线属性标志(0:无下划线，非0:有下划线)
    'byte cstrikeout,//删除线属性标志(0:无删除线，非0:有删除线)
    'byte ncharset, //字符集标识0:ansi字符集，1:系统缺省字符集
    'byte noutprecision, // 输出精度
    'byte nclipprecision, // 剪切精度
    'byte nquality, // 输出品质
    'byte npitchandfamily, // 字符间距
    'lpctstr lpszfacename // 现有系统truetype字体名称
    ')
    'Dim hdc As Long
    'hdc = g_MainSurface.DD_Surface.GetDC
    'SelectObject hdc, hfont
    'SetTextColor hdc, vbWhite
    'SetBkMode hdc, 1
    'TextOut hdc, 0, 60, "请按", 4
    'g_MainSurface.DD_Surface.ReleaseDC hdc
    With m_Font
        .size = 9
        .Weight = 400
        .Name = "黑体"
        .Bold = False
        .Bold = False
        .Charset = 0
        .Italic = 0
        .Strikethrough = 0
        .Underline = 0
    End With
    InitFont = True
    SetScreenFont m_Font
End Function

Private Function CheckActived() As Boolean
    CheckActived = True
    If g_Hwnd - GetActiveWindow Then
        CheckActived = False
    End If
End Function
