Attribute VB_Name = "xGraph"
'impactX Game Engine
'本类模块用于处理DX设备和几何绘图
'使用本类模块必须遵守:
'你可以免费使用本引擎及代码
'使用本引擎后的责任由使用者承担
'你可以任意拷贝本引擎代码，但必须保证其完整性
'希望我能得到你使用本引擎制作出的程序
'Davy.xu sunicdavy@sina.com qq:20998333
Option Explicit
Dim dx As DirectX8
Dim D3D As Direct3D8
Public D3DDevice As Direct3DDevice8
Public D3DX As D3DX8
Dim D3DWindow As D3DPRESENT_PARAMETERS '显示模式的各种参数
Const TLFVF = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1 Or D3DFVF_SPECULAR

'Transforme and Lit结构
Private Type TLVERTEX
  X As Single
  Y As Single
  Z As Single '2D渲染时不使用,设为0
  rhw As Single '不使用,设为1
  Color As Long '顶点颜色,AlphaBlend的基础颜色
  Specular As Long '光照,不使用
  tu As Single '贴图坐标(0~1) x轴
  tv As Single '贴图坐标(0~1) y轴
End Type

'图形处理方式
Enum ENUM_XG_PROCESS
    xgSOFTWARE = D3DCREATE_SOFTWARE_VERTEXPROCESSING '软件模拟
    xgHARDWARE = D3DCREATE_HARDWARE_VERTEXPROCESSING '硬件模拟
    xgPUREDEVICE = D3DCREATE_PUREDEVICE '纯硬件模拟
    xgAUTO = 0 '自动识别(缺省)
End Enum
'不带透明的颜色
Enum ENUM_XG_COLOR
    xgBLACK = &HFF000000 '黑
    xgWHITE = &HFFFFFFFF '白
    
    xgRED = &HFFFF0000 '红
    xgGREEN = &HFF00FF00 '绿
    xgBLUE = &HFF0000FF '蓝
    
    xgYELLOW = &HFFFFFF00 '黄
    xgMAGENTA = &HFFFF00FF '洋红
    xgCYAN = &HFFFF00FF '青色
End Enum
Enum ENUM_DISPLAYMODE
    xgWindow = 1
    xgFullScreen = 0
End Enum

Enum ENUM_DEVICESTATE
    xgDeviceOK = 0
    xgDeviceLost = D3DERR_DEVICELOST
    xgDeviceNotReset = D3DERR_DEVICENOTRESET
End Enum
'字体
Dim xgMainFont As D3DXFont

    
'Fps 显示
Private fpsTimer As Long
Private cCount As Integer
Private cFPS As Integer
'时间速度控制(LimitFPS)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private m_LastTime As Long
'功能:初始化DirectGraph
'参数:全屏模式下的水平和垂直分辨率,显示窗口的句柄,D3D处理方式
'注意:初始化后,屏幕的颜色数被强制调节为16位色
Public Function InitDXGraph(ByVal ResWidth As Integer, ByVal ResHeight As Integer, ByVal hWnd As Long, Optional DisplayMode As ENUM_DISPLAYMODE = xgWindow, Optional Process As ENUM_XG_PROCESS) As Boolean
    On Error GoTo ErrH
    Dim BehaviorFlag As Long
    Dim caps As D3DCAPS8
    InitDXGraph = False
    Set dx = New DirectX8
    '呼叫设定D3D
    Set D3D = dx.Direct3DCreate()
    
    Set D3DX = New D3DX8
    
    Dim DMode As D3DDISPLAYMODE
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DMode
    If DisplayMode = xgWindow Then
        D3DWindow.Windowed = 1
        D3DWindow.SwapEffect = D3DSWAPEFFECT_DISCARD 'D3DSWAPEFFECT_COPY_VSYNC '
        D3DWindow.BackBufferFormat = DMode.Format
    Else
        '页交换(Frontbuffer及Backbuffer)功能的选择,我们要使用全屏幕模式=D3DSWAPEFFECT_FLIP.
        D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
        '背景頁(Backbuffer)的数量
        D3DWindow.BackBufferCount = 1
        '背景頁的單位圖素格式,D3DFMT_R5G6B5=16bits,65536色,高彩,
        '當然您也可以把它設定為全彩(D3DFMT_R8G8B8).而我們再設定背景頁時其實就是在決定我們的顯示模式
        D3DWindow.BackBufferFormat = D3DFMT_R5G6B5
            D3DWindow.EnableAutoDepthStencil = 1
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
        '水平垂直分辨率
        D3DWindow.BackBufferWidth = ResWidth
        D3DWindow.BackBufferHeight = ResHeight
    End If
    
    '作用window的hWnd
    D3DWindow.hDeviceWindow = hWnd
    If Process = xgAUTO Then
            D3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, caps
            '自动选择合适的硬件显示设备
            If caps.DevCaps And D3DDEVCAPS_PUREDEVICE Then
               BehaviorFlag = D3DCREATE_PUREDEVICE Or D3DCREATE_HARDWARE_VERTEXPROCESSING
            Else
               If caps.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT Then
                  BehaviorFlag = D3DCREATE_HARDWARE_VERTEXPROCESSING
               Else
                  BehaviorFlag = D3DCREATE_SOFTWARE_VERTEXPROCESSING
               End If
            End If
    Else
            BehaviorFlag = Process
    End If
    
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, BehaviorFlag, D3DWindow)
    
    If Err.Number Then
        Err.Clear
        'Try to create a reference device
        Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, _
        D3DDEVTYPE_REF, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
        D3DWindow)
        'If that too fails, return error and quit
    End If
    
    
    D3DDevice.SetVertexShader TLFVF '告訴D3D描繪方式使用我們設定的方式
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE '贴图时开启ALPHA通道(AlphaBlend使用)
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA '渲染时开启透明色
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA '同上
    
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
 SetTextType "SYSTEM", 11
InitDXGraph = True
 Exit Function
ErrH:
    Debug.Print "Err [InitDXGraph] 初始化DX错误"
    Debug.Print "错误可能原因:显卡驱动不正确,显卡不支持Direct3D,禁用Direct3D,显存不够"
    MsgBox "错误可能原因:显卡驱动不正确,显卡不支持Direct3D,禁用Direct3D,显存不够" & vbNewLine & "DirectX错误:" & Hex(Err.Number), vbCritical, "初始化DX错误"
    End
End Function
'功能:给屏幕上底色
'参数:用以涂抹屏幕的颜色
Public Sub PaintScreen(ByVal BackColor As ENUM_XG_COLOR)
On Error GoTo ErrH
    If D3DDevice Is Nothing Then End
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, BackColor, 1#, 0
    Exit Sub
ErrH:
    Debug.Print "Err: [PaintScreen] 函数调用错误"
    End
End Sub
'功能:清除缓冲,准备设备
Public Sub RenderBegin()
On Error GoTo ErrH
    If D3DDevice Is Nothing Then End
    D3DDevice.BeginScene
    Exit Sub
ErrH:
    Debug.Print "Err: [RenderBegin] 函数调用错误,RenderBegin和RenderEnd必须配对使用"
    End
End Sub
'功能:结束渲染,计算FPS,缓冲区图形翻转到主页面
Public Sub RenderEnd(Optional FrmInteval As Long = 1000)
On Error Resume Next
    If D3DDevice Is Nothing Then Exit Sub
    D3DDevice.EndScene '结束渲染
    '计算FPS
    If timeGetTime - fpsTimer > FrmInteval Then
        cFPS = cCount
        cCount = 0
        fpsTimer = timeGetTime
    Else
        cCount = cCount + 1
    End If
    D3DDevice.Present ByVal 0, ByVal 0, ByVal 0, ByVal 0
    Exit Sub
ErrH:
    Debug.Print "Err: [RenderEnd] 函数调用错误,RenderBegin和RenderEnd必须配对使用"
    End
End Sub
'功能:限制FPS
'参数:需要的FPS值
Public Sub LimitFPS(Frame As Integer)
'帧每秒
    Do Until timeGetTime - m_LastTime > 1000 / Frame: Loop
    m_LastTime = timeGetTime
End Sub
'功能:开始文字渲染
'注意:请尽量将输出的文字放在BeginText和EndText之间
Public Sub BeginText()
    xgMainFont.Begin
End Sub
'功能:结束文字渲染
Public Sub EndText()
    xgMainFont.End
End Sub
'功能:在屏幕上绘制文字
'参数:文字,坐标,显示颜色
Public Sub DrawText(ByVal sText As String, X As Integer, Y As Integer, Optional Color As ENUM_XG_COLOR = xgWHITE)
On Error GoTo ErrH
    Dim rcText As RECT
    xgMainFont.Begin
    With rcText
        .Left = X
        .Top = Y
    End With
    xgMainFont.DrawTextW sText, -1, rcText, 0, Color
    xgMainFont.End
    Exit Sub
ErrH:
    Debug.Print "Err [DrawText] 绘制字体时错误"
    End
End Sub
'功能:设置显示文字
'参数:字体名字(在Word等软件里可以找到),字体大小
Public Sub SetTextType(Name As String, Optional Size As Integer = 11)
On Error GoTo ErrH
    Dim xgFontDesc As IFont
    Dim xgFont As New StdFont
    Set xgMainFont = Nothing
    xgFont.Name = Name ' "Times New Roman"
    xgFont.Size = Size '8
    xgFont.Bold = True
    Set xgFontDesc = xgFont
    Set xgMainFont = D3DX.CreateFont(D3DDevice, xgFontDesc.hFont)
    Exit Sub
ErrH:
    Debug.Print "Err [SetTextType] 创建字体时错误"
    End
End Sub
'D3D内部使用,构建一个TLV顶点结构
Private Function CreateTLVertex(X As Single, Y As Single, ByVal Color As Long, tu As Single, tv As Single) As TLVERTEX
   CreateTLVertex.X = X
   CreateTLVertex.Y = Y
   CreateTLVertex.Z = 0
   CreateTLVertex.rhw = 1
   CreateTLVertex.Color = Color
   CreateTLVertex.Specular = 0
   CreateTLVertex.tu = tu
   CreateTLVertex.tv = tv
End Function
'功能:画一条线
'参数:线条的两个点坐标,颜色
Public Sub DrawLine(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, Color As ENUM_XG_COLOR)
On Error GoTo ErrH
    Dim tVer(0 To 1) As TLVERTEX
    Dim BlankTexture As Direct3DTexture8
    tVer(0) = CreateTLVertex(CSng(X1), CSng(Y1), Color, 0, 0)
    tVer(1) = CreateTLVertex(CSng(X2), CSng(Y2), Color, 0, 0)

    D3DDevice.SetTexture 0, BlankTexture
    D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 2, tVer(0), Len(tVer(0))
    
    Exit Sub
ErrH:
    Debug.Print "Err [DrawLine] 绘图时错误"
    End
End Sub
Public Sub DrawPoint(X As Integer, Y As Integer, Color As ENUM_XG_COLOR)
On Error GoTo ErrH
    Dim tVer As TLVERTEX
    Dim BlankTexture As Direct3DTexture8
    tVer = CreateTLVertex(CSng(X), CSng(Y), Color, 0, 0)
    D3DDevice.SetTexture 0, BlankTexture
    D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, 2, tVer, Len(tVer)
    Exit Sub
ErrH:
    Debug.Print "Err [DrawPoint] 绘图时错误"
    End
End Sub

'画圆
Public Sub DrawCircle(X As Integer, Y As Integer, Radius As Integer, Color As ENUM_XG_COLOR)
    Dim iX As Integer, iY As Integer
    Static LastX As Integer, LastY As Integer
    Dim Angle As Single
    Dim tVer(0 To 1) As TLVERTEX
    Dim BlankTexture As Direct3DTexture8

    For Angle = 0 To 2 * 3.14 Step 3.14 / 30 '决定圆的圆度，step值越大，越不圆
        iX = X + (Radius * Cos(Angle))
        iY = Y + (Radius * Sin(Angle))
        If Not (LastX = 0 And LastY = 0) Then
            tVer(0) = CreateTLVertex(CSng(iX), CSng(iY), Color, 0, 0)
            tVer(1) = CreateTLVertex(CSng(LastX), CSng(LastY), Color, 0, 0)
            D3DDevice.SetTexture 0, BlankTexture
            D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 2, tVer(0), Len(tVer(0))
        End If
        LastX = iX
        LastY = iY
    Next Angle
End Sub

'功能:画一个矩形
'参数:矩形的四个角,颜色
Public Sub DrawRect(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, Color As ENUM_XG_COLOR)
    DrawLine Left, Top, Right, Top, Color
    DrawLine Right, Top, Right, Bottom, Color
    DrawLine Left, Bottom, Right, Bottom, Color
    DrawLine Left, Top, Left, Bottom, Color
End Sub
'功能:画一个矩形并填充
'参数:矩形的四个角,颜色
Public Sub DrawRectFill(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, Color As ENUM_XG_COLOR)
    On Error GoTo ErrH
    Dim tVer(0 To 3) As TLVERTEX
    Dim BlankTexture As Direct3DTexture8
    Dim t As Integer
    '颠倒处理
    If Bottom < Top Then t = Bottom: Bottom = Top: Top = t
    If Right < Left Then t = Right: Right = Left: Left = t
    tVer(0) = CreateTLVertex(CSng(Left), CSng(Top), Color, 0, 0)
    tVer(1) = CreateTLVertex(CSng(Right), CSng(Top), Color, 0, 0)
    tVer(2) = CreateTLVertex(CSng(Left), CSng(Bottom), Color, 0, 0)
    tVer(3) = CreateTLVertex(CSng(Right), CSng(Bottom), Color, 0, 0)
    
    D3DDevice.SetTexture 0, BlankTexture
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tVer(0), Len(tVer(0))
    Exit Sub
ErrH:
    Debug.Print "Err [DrawRectFill] 绘图时错误"
    End
End Sub
'功能:画一个4色渐变矩形并填充
'参数:矩形的四个角,四个角颜色
Public Sub DrawRectGradual(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, _
Color_LeftTop As ENUM_XG_COLOR, Color_LeftRight As ENUM_XG_COLOR, Color_BottomLeft As ENUM_XG_COLOR, Color_BottomRight As ENUM_XG_COLOR)
    On Error GoTo ErrH
    Dim tVer(0 To 3) As TLVERTEX
    Dim BlankTexture As Direct3DTexture8
    Dim t As Integer
    '颠倒处理
    If Bottom < Top Then t = Bottom: Bottom = Top: Top = t
    If Right < Left Then t = Right: Right = Left: Left = t
    tVer(0) = CreateTLVertex(CSng(Left), CSng(Top), Color_LeftTop, 0, 0)
    tVer(1) = CreateTLVertex(CSng(Right), CSng(Top), Color_LeftRight, 0, 0)
    tVer(2) = CreateTLVertex(CSng(Left), CSng(Bottom), Color_BottomLeft, 0, 0)
    tVer(3) = CreateTLVertex(CSng(Right), CSng(Bottom), Color_BottomRight, 0, 0)
    
    D3DDevice.SetTexture 0, BlankTexture
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, tVer(0), Len(tVer(0))
    Exit Sub
ErrH:
    Debug.Print "Err [DrawRectFill] 绘图时错误"
    End
End Sub

'功能:卸载DirectGraph
Public Sub UnloadDXGraph()
    Set D3DX = Nothing
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set dx = Nothing
End Sub

Public Function GetDeviceState() As ENUM_DEVICESTATE
    GetDeviceState = D3DDevice.TestCooperativeLevel
End Function

Public Sub DeviceReset()
    D3DDevice.Reset D3DWindow
    D3DDevice.SetVertexShader TLFVF '告訴D3D描繪方式使用我們設定的方式
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE '贴图时开启ALPHA通道(AlphaBlend使用)
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA '渲染时开启透明色
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA '同上
    
End Sub
'获得FPS
Public Property Get GetFPS() As Integer
    GetFPS = cFPS
End Property

