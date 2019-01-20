VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===========================Class Screen============================
'       GameEngine中的显示接口部分
'Create(Width,Height,colornum,windowed):
'   创建一个视口Width×Height，颜色数为colornum，windowed窗体或全屏，
'   若创建成功，返回一个DirectDrawSurface7的缓冲表面。
'Render:
'   将缓冲表面的数据刷新到主表面显示出来。
'Destrory:
'   销毁该视口。
'm_MainSurface:主表面。
'm_BufferSurface:缓冲表面。
'=====================================================================
Public RefreshDelay As Long
Public FPS As Long

Private m_Windowed As Boolean
Private m_ViewHwnd As Long
Private m_ScreenWidth As Long
Private m_ScreenHeight As Long
Private m_ViewWidth As Long
Private m_ViewHeight As Long

Private m_LastPaintTime As LARGE_INTEGER
Private m_BaseFrequency As LARGE_INTEGER
Private m_BaseDethaTime As Double

Private m_MainSurface As DirectDrawSurface7
Private m_BufferSurface As DirectDrawSurface7
Private m_GammaControl As DirectDrawGammaControl
Private m_OldGammamp As DDGAMMARAMP

Private m_LightTable() As Byte
Private m_LightTableUsed As Boolean
Private m_lpLightTable As Long                  '光照表首地址]

Private m_ZBuffer() As Integer                  '16bits Z_Buffer
Private m_lpZBuffer As Long                     'Z_Buffer表首地址
Private m_ZBufferEnabled As Boolean

Private m_BltCmds() As TStdBlt                  '伪作图指令Buffer
Private m_BltCmdNum As Long
Private Const BLT_BUFFER_LENTH = 128
Private m_StdBltQword As Long
Private m_BltCmdEnabled As Boolean              '

'---------------------------------------------------------------------
'设置屏幕颜色数API声明
'---------------------------------------------------------------------
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal HDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As size) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal HDC As Long, ByVal nIndex As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long

Private Const CDS_UPDATEREGISTRY = &H1, CDS_TEST = &H4, DISP_CHANGE_SUCCESSFUL = 0
Private Const EWX_FORCE = 4, CCDEVICENAME = 32, CCFORMNAME = 32, DISP_CHANGE_RESTART = 1

Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const DM_DISPLAYFREQUENCY = &H400000

Private Const WM_DISPLAYCHANGE = &H7E&
Private Const HWND_BROADCAST = &HFFFF&
Private Const SPI_SETNONCLIENTMETRICS = 42

Private Type size
    cx As Long
    cy As Long
End Type

Private Type DEVMODE
    dmDeviceName       As String * 32
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * 32
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type
'DDGAMMARAMP

Public Sub Destroy()
    Set m_BufferSurface = Nothing
    Set m_MainSurface = Nothing
    Set m_GammaControl = Nothing
End Sub

Private Sub Class_Initialize()
    m_Windowed = True
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Function DrawSurface() As DirectDrawSurface7
    Set DrawSurface = m_MainSurface
End Function

Public Function Create(hWnd As Long, Width As Long, Height As Long, Optional colorNum As Long = 16, _
                       Optional windowed As Boolean = True, Optional SYSTEMMEMORY As Boolean = False) As DirectDrawSurface7
    '返回m_BufferSurface
    Dim ddsd As DDSURFACEDESC2
    Dim ddCliper As DirectDrawClipper
    On Error GoTo errDDrawFaild

    g_ViewRect.Right = Width
    g_ViewRect.Bottom = Height
    m_ViewWidth = Width
    m_ViewHeight = Height
    '创建光照表
    m_ViewHwnd = hWnd
    m_Windowed = windowed
    '检查窗体模式下颜色是否１６ｂｉｔｓ
    If m_Windowed Then
        If Not SetupDisplayMode Then
            Set Create = Nothing
            MsgBox "显示模式设置失败！请手工设置系统16位色彩模式！", , "VBGameEngine"
            Err.Raise vbObjectError + 513, "VBGameEngine", "SetupDisplayMode Failed!"
            Exit Function
        End If
    End If
    '创建DirectDraw,设置显示模式合作层
    Set g_DD = g_DX7.DirectDrawCreate("")
    If g_DD Is Nothing Then Exit Function

    If m_Windowed Then
        g_DD.SetCooperativeLevel m_ViewHwnd, DDSCL_NORMAL
    Else
        g_DD.SetCooperativeLevel m_ViewHwnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX
        g_DD.SetDisplayMode Width, Height, colorNum, 0, DDSDM_DEFAULT
    End If
    '创建主表面
    ddsd.lFlags = DDSD_CAPS
    ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set m_MainSurface = g_DD.CreateSurface(ddsd)
    '创建缓冲表面
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If SYSTEMMEMORY Then
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Else
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    End If
    ddsd.lWidth = Width
    ddsd.lHeight = Height
    Set m_BufferSurface = g_DD.CreateSurface(ddsd)

    m_BufferSurface.SetForeColor vbWhite
    '设置剪切
    Set ddCliper = g_DD.CreateClipper(0)
    ddCliper.SetHWnd m_ViewHwnd
    m_MainSurface.SetClipper ddCliper
    '刷新象素格式（含字体单位象素大小）
    RefreshPixelFormat
    '得到主表面的m_GammaControl
    Set m_GammaControl = m_MainSurface.GetDirectDrawGammaControl
    m_GammaControl.GetGammaRamp DDSGR_DEFAULT, m_OldGammamp

    Set Create = m_BufferSurface

    '开启Fps控制
    QueryPerformanceFrequency m_BaseFrequency
    If m_BaseFrequency.low < 0 Then
        m_BaseDethaTime = 1000 / CDbl(2147483648# - m_BaseFrequency.low)
    Else
        m_BaseDethaTime = 1000 / m_BaseFrequency.low
    End If
    QueryPerformanceCounter m_LastPaintTime
    
    QueryPerformanceCounter m_LastPaintTime

Exit Function

errDDrawFaild:
    MsgBox Err.Description, , "VBGameEngine"
    Set Create = Nothing
End Function

Public Function TimeOut() As Boolean
    Dim newTime As LARGE_INTEGER
    Dim DethaTime As Long
    On Error GoTo errhand
    QueryPerformanceCounter newTime
    DethaTime = (newTime.low - m_LastPaintTime.low) * m_BaseDethaTime
    TimeOut = (DethaTime >= RefreshDelay)

    Exit Function

errhand:
    TimeOut = True
End Function

Public Function CreateLightTable() As Long
    '创建光照表
    '目前只有在全屏模式支持光照表
    Debug.Assert m_Windowed = False

    m_LightTableUsed = True
    If m_lpLightTable <> 0 Then CreateLightTable = m_lpLightTable: Exit Function

    ReDim m_LightTable(m_ViewWidth - 1, m_ViewHeight - 1)
    m_lpLightTable = VarPtr(m_LightTable(0, 0))
    CreateLightTable = m_lpLightTable
    SetAmbientLight 255
End Function

Public Sub CloseLightTable()
    '关闭光线表
    m_LightTableUsed = False
End Sub

Public Sub SetAmbientLight(ByVal Light As Byte, Optional UseGammamped As Boolean = False)
    '设置环境光强
    Dim NewGammamp As DDGAMMARAMP, i As Long, j As Long
    If Not m_Windowed And UseGammamped Then
        '优先控制！
        For i = 0 To 255
            j = CLng(i) * Light
            If j > 32767 Then j = j - 65536
            NewGammamp.red(i) = j
            NewGammamp.green(i) = j
            NewGammamp.blue(i) = j
        Next i
        m_GammaControl.SetGammaRamp DDSGR_DEFAULT, NewGammamp
        Exit Sub
    End If

    If Not m_LightTableUsed Then Exit Sub
    Dim DataBytes As Long
    DataBytes = m_ViewWidth * CLng(m_ViewHeight)
    fastmemset ByVal m_lpLightTable, DataBytes, Light
End Sub

Public Sub BltToLightTable(lightSurface As CSurface, ByVal x As Long, ByVal y As Long, Optional ByVal frame As Long = 0)
    '支持动态光照
    Dim srcRect As RECT, dstRect As RECT
    Dim swidth As Long
    If Not g_MainSurface.ClippeSurface(lightSurface, x, y, srcRect, dstRect, frame) Then Exit Sub
    swidth = dstRect.Right - dstRect.Left

    swidth = swidth And &HFFFC
    If swidth < 4 Then Exit Sub

    'Bug Here !First Must Lock
    blt_to_lighttable_mmx ByVal m_lpLightTable, dstRect.Left, dstRect.Top, _
                          swidth, dstRect.Bottom - dstRect.Top, _
                          m_ViewWidth, ByVal lightSurface.DataAddress, srcRect.Left, srcRect.Top, lightSurface.DataPitch

End Sub

'-Z_Buffer
Public Sub EnableZBuffer()
    If m_lpZBuffer = 0 Then
        ReDim m_ZBuffer(m_ViewWidth - 1, m_ViewHeight - 1)
        m_lpZBuffer = VarPtr(m_ZBuffer(0, 0))
    End If
    m_ZBufferEnabled = True
End Sub

Public Sub ClearZBuffer(Optional ByVal z As Long = &H7FFF)
    Debug.Assert m_lpZBuffer <> 0
    memsetw ByVal m_lpZBuffer, m_ViewWidth * CLng(m_ViewHeight), z
End Sub

Public Sub BltWithZBuffer(srcSurface As CSurface, ByVal x As Long, ByVal y As Long, z As Long, Optional ByVal frame As Long = 0)
    Dim srcRect As RECT, dstRect As RECT
    Dim swidth As Long
    Debug.Assert m_lpZBuffer <> 0
    
    If Not g_MainSurface.ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame) Then Exit Sub
    swidth = dstRect.Right - dstRect.Left

    swidth = swidth And &HFFFC
    If swidth < 4 Then Exit Sub

    zbuffer_blt_mmx ByVal g_MainSurface.DataAddress, ByVal m_lpZBuffer, dstRect.Left, dstRect.Top, z, _
                    m_ViewWidth, ByVal srcSurface.DataAddress, srcRect.Left, srcRect.Top, srcSurface.DataPitch, _
                    swidth, srcRect.Bottom - srcRect.Top, SYS_KEYCOLOR

End Sub

Public Function GetZBuffer(ByVal x As Long, ByVal y As Long) As Long
    Debug.Assert m_lpZBuffer <> 0 And x >= 0 And x < m_ViewWidth And y >= 0 And y < m_ViewHeight
    GetZBuffer = m_ZBuffer(x, y) And &H7FFF
End Function

Public Sub Render()
    Dim dstRect As RECT
    Dim bRestore As Boolean
    Dim ddsd As DDSURFACEDESC2
    Dim dw As Long, dh As Long
    Dim lps0 As Long, lps1 As Long, lspitch As Long, ldpitch As Long
    Dim stime As Long
    '性能计数器
    On Error Resume Next
    stime = m_LastPaintTime.low
    QueryPerformanceCounter m_LastPaintTime
    FPS = 1000 / ((m_LastPaintTime.low - stime) * m_BaseDethaTime)

    'Test Alt+tab
    Do Until ExModeActive
        DoEvents
        Sleep 1
        bRestore = True
    Loop
    If bRestore Then
        g_DD.RestoreAllSurfaces
        'InitAllSurface
        '有问题
    End If

    '绘制Blt指令集
    RenderBltCmds

    If m_Windowed Then
        Dim tp As POINTAPI
        GetDCOrgEx GetDC(m_ViewHwnd), tp
        dstRect.Left = tp.x
        dstRect.Top = tp.y
        dstRect.Right = tp.x + m_ViewWidth
        dstRect.Bottom = tp.y + m_ViewHeight
    Else
        dstRect = g_ViewRect
    End If

    If g_Mode = ALHPA_SPEED_MODE And Not m_Windowed Then
        '手工实现Blt的快速运算，暂时使用DX7的Blt
        If Not m_LightTableUsed Then
            'StickStart
            m_MainSurface.Lock dstRect, ddsd, DDLOCK_SURFACEMEMORYPTR, 0
            lps0 = ddsd.lpSurface
            ldpitch = ddsd.lPitch

            m_BufferSurface.Lock g_ViewRect, ddsd, DDLOCK_SURFACEMEMORYPTR, 0
            lps1 = ddsd.lpSurface
            lspitch = ddsd.lPitch

            dw = (dstRect.Right - dstRect.Left) And &HFFFC
            bltfast ByVal lps0, dstRect.Left, dstRect.Top, ldpitch, dw, dstRect.Bottom - dstRect.Top, _
                    ByVal lps1, 0, 0, lspitch

            m_BufferSurface.Unlock g_ViewRect
            m_MainSurface.Unlock dstRect
            'Debug.Print StickEnd; "fastblt"
        Else
            m_MainSurface.Lock dstRect, ddsd, DDLOCK_SURFACEMEMORYPTR, 0
            lps0 = ddsd.lpSurface
            ldpitch = ddsd.lPitch
            m_BufferSurface.Lock g_ViewRect, ddsd, DDLOCK_SURFACEMEMORYPTR, 0
            lps1 = ddsd.lpSurface
            lspitch = ddsd.lPitch
            dw = (dstRect.Right - dstRect.Left) And &HFFFC
            fastlight_565_mmx ByVal lps0, ByVal lps1, m_LightTable(0, 0), dstRect.Left, dstRect.Top, _
                              ldpitch, lspitch, g_ViewRect.Right, dw, dstRect.Bottom - dstRect.Top
            m_BufferSurface.Unlock g_ViewRect
            m_MainSurface.Unlock dstRect

        End If

    Else
        'StickStart
        m_MainSurface.Blt dstRect, m_BufferSurface, g_ViewRect, DDBLT_WAIT
        'Debug.Print StickEnd; "blt"
    End If

End Sub

Private Function ExModeActive() As Boolean
    If g_DD.TestCooperativeLevel = DD_OK Then ExModeActive = True
End Function

Public Sub FadeIn()
    Dim NewGammamp As DDGAMMARAMP, i As Long, j As Long, k As Long
    '全屏下淡入
    If g_Windowed Then
        For i = 255 To 200 Step -1
            Sleep 10
            g_MainSurface.DrawAlphaRect 0, 0, g_MainSurface.Width, g_MainSurface.Height, i
            Render
        Next i
        Exit Sub
    End If
    For i = 256 To 0 Step -4
        Sleep 10
        For j = 0 To 255
            k = CLng(j) * CLng(i)
            If k > 32767 Then k = k - 65536
            NewGammamp.red(j) = k
            NewGammamp.green(j) = k
            NewGammamp.blue(j) = k
        Next
        m_GammaControl.SetGammaRamp DDSGR_DEFAULT, NewGammamp
    Next
End Sub

Public Sub FadeOut()
    Dim NewGammamp As DDGAMMARAMP, i As Long, j As Long, k As Long
    '全屏下淡出 不支持窗体模式
    For i = 0 To 256 Step 4
        Sleep 10
        For j = 0 To 255
            k = CLng(j) * CLng(i)
            If k > 32767 Then k = k - 65536
            NewGammamp.red(j) = k
            NewGammamp.green(j) = k
            NewGammamp.blue(j) = k
        Next
        m_GammaControl.SetGammaRamp DDSGR_DEFAULT, NewGammamp
    Next
End Sub

Public Sub BackUpGammamp()
    m_GammaControl.SetGammaRamp DDSGR_DEFAULT, m_OldGammamp
End Sub

Public Sub RefreshPixelFormat()
    '读入主表面的象素格式
    Dim tPixelFormat As DDPIXELFORMAT
    m_MainSurface.GetPixelFormat tPixelFormat
    g_RBitsMask = tPixelFormat.lRBitMask
    g_GBitsMask = tPixelFormat.lGBitMask
    g_BBitsMask = tPixelFormat.lBBitMask
    '得到单位字的宽高
    Dim tSize As size, HDC As Long
    HDC = m_BufferSurface.GetDC
    GetTextExtentPoint32 HDC, "A", 1, tSize
    g_XPiexlsPerWord = tSize.cx
    g_YPiexlsPerWord = tSize.cy
    m_BufferSurface.ReleaseDC HDC
End Sub

Private Function SetupDisplayMode() As Boolean
    '设置窗体下的显示模式
    Dim ScreenBits As Long
    Dim lRetVal As Long, tDevMode As DEVMODE

    lRetVal = EnumDisplaySettings(0, -1, tDevMode)

    m_ScreenWidth = tDevMode.dmPelsWidth
    m_ScreenHeight = tDevMode.dmPelsHeight
    ScreenBits = tDevMode.dmBitsPerPel

    If ScreenBits = 16 Then
        SetupDisplayMode = True
    Else
        MsgBox "系统将设置为16位色！", , "GameEngine"
        tDevMode.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL Or DM_DISPLAYFREQUENCY
        tDevMode.dmPelsWidth = m_ScreenWidth
        tDevMode.dmPelsHeight = m_ScreenHeight
        tDevMode.dmBitsPerPel = 16

        lRetVal = ChangeDisplaySettings(tDevMode, CDS_UPDATEREGISTRY)
        SetupDisplayMode = (lRetVal = DISP_CHANGE_SUCCESSFUL)
    End If
End Function

'------------------------------
'以下代码执行作图指令Buffer
'------------------------------
Public Sub Blt(srcSurface As CSurface, bltmode As STD_BLT_MODE, x As Long, y As Long, _
    Optional frame As Long = 0, _
    Optional effect As Long = 0, _
    Optional k As Single = 1, _
    Optional z As Long = 0, _
    Optional ID As Long = 0)
    '使用Blt伪代码
    Debug.Assert m_BltCmdEnabled And z >= 0
    
    '调整Z的位置实现深度
    Dim AppendPos As Long
    
    AppendPos = AdjustZPos(z)
    
    m_BltCmds(AppendPos).x = x
    m_BltCmds(AppendPos).y = y
    m_BltCmds(AppendPos).z = z
    m_BltCmds(AppendPos).k = k
    m_BltCmds(AppendPos).lpSurface = ObjPtr(srcSurface)
    m_BltCmds(AppendPos).bltmode = bltmode
    m_BltCmds(AppendPos).effect = effect
    m_BltCmds(AppendPos).frame = frame
    m_BltCmds(AppendPos).ID = ID
    m_BltCmdNum = m_BltCmdNum + 1
    
    If m_BltCmdNum > UBound(m_BltCmds) Then
        ReDim Preserve m_BltCmds(UBound(m_BltCmds) + 1)
    End If
End Sub

Private Sub RenderBltCmds()
    'Blt渲染指令
    If Not m_BltCmdEnabled Then
        Exit Sub
    End If
    
    Dim i As Long, bltmode As STD_BLT_MODE, srcSurface As CSurface
    For i = m_BltCmdNum - 1 To 0 Step -1
        CopyMemory srcSurface, m_BltCmds(i).lpSurface, 4

        bltmode = m_BltCmds(i).bltmode
        If bltmode = STD_BLT Then
            g_MainSurface.Blt srcSurface, m_BltCmds(i).x, m_BltCmds(i).y, m_BltCmds(i).frame, m_BltCmds(i).k
        ElseIf bltmode = FAST_BLT Then
            g_MainSurface.FastBlt srcSurface, m_BltCmds(i).x, m_BltCmds(i).y, m_BltCmds(i).frame
        ElseIf bltmode = FAST_BLT_EX Then
            g_MainSurface.FastBltEx srcSurface, m_BltCmds(i).x, m_BltCmds(i).y, m_BltCmds(i).frame
        ElseIf bltmode = ALPHA_BLT Then
            g_MainSurface.AlphaBlt srcSurface, m_BltCmds(i).x, m_BltCmds(i).y, CByte(m_BltCmds(i).effect), m_BltCmds(i).frame
        ElseIf bltmode = ADDTIVE_BLT Then
            g_MainSurface.Additive srcSurface, m_BltCmds(i).x, m_BltCmds(i).y, CByte(m_BltCmds(i).effect), m_BltCmds(i).frame
        ElseIf bltmode = SUB_BLT Then
            g_MainSurface.Subtract srcSurface, m_BltCmds(i).x, m_BltCmds(i).y, CByte(m_BltCmds(i).effect), m_BltCmds(i).frame
        ElseIf bltmode = MASK_BLT Then
            g_MainSurface.GreenMask srcSurface, m_BltCmds(i).effect, m_BltCmds(i).x, m_BltCmds(i).y, m_BltCmds(i).frame
        ElseIf bltmode = EDGELINE_BLT Then
            g_MainSurface.BltWithEdgeline srcSurface, m_BltCmds(i).x, m_BltCmds(i).y, m_BltCmds(i).effect, m_BltCmds(i).frame
        ElseIf bltmode = ROTE_BLT Then
            g_MainSurface.RotateRect srcSurface, m_BltCmds(i).x, m_BltCmds(i).y, m_BltCmds(i).k, m_BltCmds(i).frame
        ElseIf bltmode = TEXT_BLT Then
            'g_MainSurface.TextOut
        ElseIf bltmode = ZBUFFER_BLT Then
            
        End If
        
        CopyMemory srcSurface, 0&, 4
    Next
    
    m_BltCmdNum = 0
End Sub

Public Sub EnableBltCmd(Optional state As Boolean = True)
    '开启Blt指令
    If state Then
        ReDim m_BltCmds(BLT_BUFFER_LENTH - 1)
    Else
        Erase m_BltCmds
    End If
    m_BltCmdEnabled = state
    m_StdBltQword = LenB(m_BltCmds(0))
    Debug.Assert m_StdBltQword Mod 8 = 0           '便于mmx的处理
End Sub

Private Function AdjustZPos(z As Long) As Long
    Dim i As Long
    
    If m_BltCmdNum = 0 Then
        AdjustZPos = 0
        Exit Function
    End If
    
    '首先！判断是否是特殊Z（Z<m_BltCmds(0) or Z>m_BltCmds(m_BltCmdNum)）
    If z >= m_BltCmds(m_BltCmdNum - 1).z Then
        AdjustZPos = m_BltCmdNum
        Exit Function
    ElseIf z <= m_BltCmds(0).z Then
        AdjustZPos = 0
        '快速依次移动
        ShrZPos 0
        Exit Function
    End If
    
    '执行二分法快速插入排序
    Dim zMin As Long, zMax As Long, zCenter As Long
    zMin = 0
    zMax = m_BltCmdNum
    
    Do
        zCenter = (zMin + zMax) / 2
        If z >= m_BltCmds(zCenter).z Then
            zMin = zCenter
        Else
            zMax = zCenter
        End If
    Loop While zMax - zMin > 1

    ShrZPos zMax
    AdjustZPos = zMax
End Function

Private Sub ShrZPos(start As Long)
    '右移一位
    'Use CopyMemory To Adjust Datas
    '请求反向CopyMemory的Dll支持！
    memrecopy_mmx m_BltCmds(m_BltCmdNum).lpSurface, m_BltCmds(m_BltCmdNum - 1).lpSurface, (m_BltCmdNum - start) * m_StdBltQword
End Sub

Public Function GetZScreenID(ByVal x As Long, ByVal y As Long) As Long
    '得到屏幕指定点的ZID的值
    Debug.Assert m_BltCmdEnabled
    Dim i As Long, srcSurface As CSurface
    
    For i = 0 To m_BltCmdNum - 1
        CopyMemory srcSurface, m_BltCmds(i).lpSurface, 4
        If Not srcSurface.IsTrans((x - m_BltCmds(i).x) / m_BltCmds(i).k, _
                                (y - m_BltCmds(i).y) / m_BltCmds(i).k, _
                                m_BltCmds(i).frame) Then
            GetZScreenID = m_BltCmds(i).ID
            CopyMemory srcSurface, 0&, 4
            Exit Function
        End If
        CopyMemory srcSurface, 0&, 4
    Next
    GetZScreenID = -1
End Function
