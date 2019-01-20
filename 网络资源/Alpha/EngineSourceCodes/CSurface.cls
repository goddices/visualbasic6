VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===========================Class  Surface============================
'Type SingleFrame:单帧图象的描叙，关键点＋区域
'm_Frames():记录了所有的帧数据，以便于抽取单独的帧
'Sub GetFrameKeyPoint FrameCount, x, y
'Function FrameRect(FrameCount)
'默认透明色：RGB(255,0,255)
'=====================================================================

Public Enum effect
    Alpha_Effect = 1
    Additive_Effect = 2
End Enum
Private Type SingleFrame
    Area As RECT
    KeyPointX As Integer
    KeyPointY As Integer
End Type
Private Type DirtyRect
    dstRect As RECT
    srcRect As RECT
    srcSurface As CSurface
End Type

Private m_Surface As DirectDrawSurface7
Private m_width As Long
Private m_height As Long
Private m_ForeColor As Long
Private m_ClipperRect As RECT
Private m_lpData  As Long                   '数据指针
Private m_lPitch As Long                    'Data Pitch
Private m_Frames() As SingleFrame           '支持动画帧的设置
Private m_Locked As Boolean
Private m_EnableRecordRect As Boolean       '是否开启区域记录
Private m_DirtyRect As RECT

Public Sub Destroy()
    ReDim m_Frames(0)
    Set m_Surface = Nothing
End Sub

Private Sub Class_Initialize()
    ReDim m_Frames(0)
    m_ForeColor = vbWhite
    m_EnableRecordRect = False
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub
'--------------------------------testing---------------------------------------------
'--------------------------------       ---------------------------------------------
Public Sub AlphaBltTest(srcSurface As CSurface, x As Long, y As Long, ByVal alph As Byte, Optional frame As Long = 0)
    Dim srcRect As RECT, dstRect As RECT
    Dim swidth As Long, sheight As Long
    On Error GoTo errhand:
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, 1) Then Exit Sub
    swidth = srcRect.Right - srcRect.Left
    sheight = srcRect.Bottom - srcRect.Top
    
    Me.BeginLock
    srcSurface.BeginLock
    swidth = swidth And &HFFFC
    swidth = srcSurface.Width And &HFFFC
    If swidth < 4 Then Exit Sub
    
    Dim t As Byte
    If Rnd > 0.95 Then t = 50 * Sin(Timer * 200) + 50
    addcolor_565_mmx ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, _
        swidth, srcSurface.Height, RGBtoDDColor(t, t, t)
    srcSurface.EndLock
    Me.EndLock
    Exit Sub
errhand:
    Debug.Print Err.Description
    'colorblend_565_mmx
End Sub

Public Property Get Width() As Long
    Width = m_width
End Property

Public Property Get Height() As Long
    Height = m_height
End Property

Public Sub ReSize(ByVal Width As Long, ByVal Height As Long)
    m_width = Width
    m_height = Height
End Sub

Public Property Get DD_Surface() As DirectDrawSurface7
    Set DD_Surface = m_Surface
End Property

Public Property Let DD_Surface(New_DD_Surface As DirectDrawSurface7)
    If New_DD_Surface Is Nothing Then Exit Property
    Set m_Surface = New_DD_Surface
    
    SetKeyColor SYS_KEYCOLOR
    SetClipper 0, 0, m_width, m_height
    
    If g_Mode = ALHPA_SPEED_MODE Then
        'Fill Data address and pitch,if in ALPHA_MODE
        Me.BeginLock
        Me.EndLock
    End If
End Property

Public Sub SetKeyColor(ByVal newKeycolor As Long)
    Dim ColorKey As DDCOLORKEY
    ColorKey.low = newKeycolor
    ColorKey.high = newKeycolor
    m_Surface.SetColorKey DDCKEY_SRCBLT, ColorKey
    
End Sub

Public Property Get DataAddress() As Long
    '得到数据区的首地址
    DataAddress = m_lpData
End Property

Public Property Get DataPitch() As Long
    DataPitch = m_lPitch
End Property

Public Property Get FrameCount() As Long
    '得到当前的桢数
    FrameCount = UBound(m_Frames()) - 1
End Property

Public Property Get EnableRecordRect() As Boolean
    EnableRecordRect = m_EnableRecordRect
End Property

Public Property Let EnableRecordRect(ByVal vNewValue As Boolean)
    m_EnableRecordRect = vNewValue
End Property

Public Function GetDirtyRect() As RECT
    GetDirtyRect = m_DirtyRect
End Function

Public Function Create(Width As Long, Height As Long, Optional SYSTEMMEMORY As Boolean = True) As Boolean
    Dim ddsd As DDSURFACEDESC2
    If g_DD Is Nothing Then Exit Function
    If Not m_Surface Is Nothing Then Set m_Surface = Nothing
    
    m_width = Width
    m_height = Height
    
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd.lWidth = Width
    ddsd.lHeight = Height
    If g_Mode = 0 Then
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    DD_Surface = g_DD.CreateSurface(ddsd)
    ReadFrames "Create"
    Create = True
    
    If g_Mode = ALHPA_SPEED_MODE Then
        'Fill Data address and pitch,if in mode 1:ALPHA_MODE
        Me.BeginLock
        Me.EndLock
    End If
End Function

Public Sub SetClipper(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    ' 注意，不要越界
    m_ClipperRect.Left = Left
    m_ClipperRect.Right = Right
    m_ClipperRect.Top = Top
    m_ClipperRect.Bottom = Bottom
End Sub

Public Function GetClipper() As RECT
    GetClipper = m_ClipperRect
End Function

Public Sub ColorFill(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)
    m_Surface.BltColorFill m_ClipperRect, RGBtoDDColor(R, G, B)
End Sub

Public Sub PutPixel(x As Long, y As Long, color As Long, Optional frame As Long = 0)
    Dim srcRect As RECT
    srcRect = Me.FrameRect(frame)
    x = x + srcRect.Left
    y = y + srcRect.Top
    If x < 0 Or y < 0 Or x > m_ClipperRect.Right - 1 Or y > m_ClipperRect.Bottom - 1 Then Exit Sub
    Me.BeginLock
    setpixel ByVal m_lpData, x, y, Me.DataPitch, color
    Me.EndLock
End Sub

Public Function Pixels(ByVal x As Long, ByVal y As Long, Optional frame As Long = 0) As Long
    '得到16Bits的象素值,其中X,Y为相对于当前Frame的位置
    Dim srcRect As RECT
    srcRect = Me.FrameRect(frame)
    x = x + srcRect.Left
    y = y + srcRect.Top
    If x < 0 Or y < 0 Or x > m_ClipperRect.Right - 1 Or y > m_ClipperRect.Bottom - 1 Then
        Pixels = SYS_KEYCOLOR
        Exit Function
    End If
    Me.BeginLock
    Pixels = GetPixel(ByVal m_lpData, x, y, Me.DataPitch)
    Me.EndLock
End Function

Public Function IsTrans(ByVal x As Long, ByVal y As Long, _
    Optional frame As Long = 0) As Boolean
    '是否该点透明,注意(X,Y)相对于关键点而非左上角
    Dim keyx As Long, keyy As Long
    Me.GetFrameKeyPoint frame, keyx, keyy

    x = x + keyx
    y = y + keyy
    
    If x < 0 Or y < 0 Or x > m_ClipperRect.Right - 1 Or y > m_ClipperRect.Bottom - 1 Or _
        Not CheckRect(Me.FrameRect(frame), x, y) Then
        IsTrans = True
        Exit Function
    End If
    Me.BeginLock
    '63519 OK &hf81f bug（&hf81f＝63519）
    'If getpixel(ByVal m_lpData, x, y, Me.DataPitch) = &HF81F Then IsTrans = True
    If GetPixel(ByVal m_lpData, x, y, Me.DataPitch) = SYS_KEYCOLOR Then IsTrans = True
    
    
    Me.EndLock
End Function

Public Sub TextOut(ByVal x As Long, ByVal y As Long, ByVal text As String, Optional color As Long = vbWhite)
    If text = "" Then Exit Sub
    On Error Resume Next
    m_Surface.SetForeColor m_Surface.GetFontBackColor
    m_Surface.DrawText x + 2, y + 2, text, False
    m_Surface.SetForeColor color
    m_Surface.DrawText x, y, text, False
    m_Surface.SetForeColor m_ForeColor          'Save Back
End Sub

Public Function LoadBMP(FileName As String, Width As Long, Height As Long) As Boolean
    Dim ddsd As DDSURFACEDESC2
    On Error GoTo errhand
    If g_DD Is Nothing Then Exit Function
    
    m_width = Width
    m_height = Height
    
    '首先创建一个DirectDrawSurface
    'If Not Create(m_width, m_height) Then LoadBMP = False: Exit Function
    
    ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd.lWidth = Width
    ddsd.lHeight = Height
    If g_Mode = 0 Then
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    
    DD_Surface = g_DD.CreateSurfaceFromFile(FileName, ddsd)
    
    If Not m_Surface Is Nothing Then
        LoadBMP = True
    Else
        Exit Function
    End If
    
    If g_Mode = ALHPA_SPEED_MODE Then
        'Fill Data address and pitch,if in mode 1:ALPHA_MODE
        Me.BeginLock
        Me.EndLock
    End If
    '载入初始区域，自动检测ReadFrames（FileName）
    LoadBMP = ReadFrames(FileName)
errhand:
End Function

Public Function LoadJPG(FileName As String) As Boolean
    Dim ddsd As DDSURFACEDESC2
    Dim tJPG As New CGDISurface
    Dim thdc As Long
    On Error GoTo errhand
    If g_DD Is Nothing Then Exit Function
    
    tJPG.Read FileName
    
    m_width = tJPG.Width
    m_height = tJPG.Height
    
    '首先创建一个DirectDrawSurface
    If Not Create(m_width, m_height) Then LoadJPG = False: Exit Function
    If Not m_Surface Is Nothing Then LoadJPG = True Else Exit Function
    
    '将GDISurface.Render->DirectDrawSurface1
    thdc = m_Surface.GetDC
    tJPG.Render thdc
    DD_Surface.ReleaseDC thdc
    'Desdroy tJPG
    Set tJPG = Nothing
    '载入初始区域，自动检测ReadFrames（FileName）
    LoadJPG = ReadFrames(FileName)
errhand:
End Function
  
Public Function LoadRes(SrcID As Long, Width As Long, Height As Long) As Boolean

End Function

'--------------------------- BeginLock ---------------------
'        锁定该DirectDrawSurface,以便象素处理
'-----------------------------------------------------------
Public Function BeginLock() As Boolean
    Dim ddsd As DDSURFACEDESC2
    If m_Locked Then BeginLock = True: Exit Function
    m_Surface.Lock m_ClipperRect, ddsd, DDLOCK_SURFACEMEMORYPTR, 0
    m_lpData = ddsd.lpSurface
    m_lPitch = ddsd.lPitch
    BeginLock = m_lpData                    'm_lpData=0 failed!
    m_Locked = True
End Function

'--------------------------- BeginLock ---------------------
'        解锁DirectDrawSurface,以便Blt操作
'-----------------------------------------------------------
Public Sub EndLock()
    If Not m_Locked Then Exit Sub
    m_Surface.Unlock m_ClipperRect
    m_Locked = False
End Sub


'------------------------------ClippeSurface--------------------------------
'自动裁剪表面,以免Blt操作越界
'SrcSurface       待裁剪表面
'x,y              待显示位置
'srcRect,dstRect  函数将处理结果填入该变量
'Frame            待显示帧数
'k                缩放系数
'----------------------------------------------------------------------------
Public Function ClippeSurface(srcSurface As CSurface, x As Long, y As Long, srcRect As RECT, dstRect As RECT, _
    Optional frame As Long = 0, Optional k As Single = 1) As Boolean
    '完成表面的裁剪
    Dim swidth As Long, sheight As Long, keyx As Long, keyy As Long
    srcSurface.GetFrameKeyPoint frame, keyx, keyy
    srcRect = srcSurface.FrameRect(frame)
    swidth = srcRect.Right - srcRect.Left
    sheight = srcRect.Bottom - srcRect.Top
    
    dstRect.Left = x - (keyx - srcRect.Left) * k
    dstRect.Top = y - (keyy - srcRect.Top) * k
    dstRect.Right = dstRect.Left + swidth * k
    dstRect.Bottom = dstRect.Top + sheight * k
    
    'Clipper
    ClippeSurface = ClipperRect(srcRect, dstRect, k)
End Function

'---------------------------------ClipperRect------------------------------
'处理区域的自动裁剪,以免Blt操作越界,返回False表明无法裁剪
'srcRect         源矩形区域
'dstRect         目标矩形区域
'k               缩放系数 k*srcRect->dstRect
'--------------------------------------------------------------------------
Private Function ClipperRect(srcRect As RECT, dstRect As RECT, Optional k As Single = 1) As Boolean
    '裁剪区域的自动处理
    If dstRect.Left < m_ClipperRect.Left Then
        srcRect.Left = srcRect.Left + (m_ClipperRect.Left - dstRect.Left) / k
        dstRect.Left = m_ClipperRect.Left
    End If
    If dstRect.Top < m_ClipperRect.Top Then
        srcRect.Top = srcRect.Top + (m_ClipperRect.Top - dstRect.Top) / k
        dstRect.Top = m_ClipperRect.Top
    End If
    If dstRect.Right > m_ClipperRect.Right Then
        srcRect.Right = srcRect.Right + (m_ClipperRect.Right - dstRect.Right) / k
        dstRect.Right = m_ClipperRect.Right
    End If
    If dstRect.Bottom > m_ClipperRect.Bottom Then
        srcRect.Bottom = srcRect.Bottom + (m_ClipperRect.Bottom - dstRect.Bottom) / k
        dstRect.Bottom = m_ClipperRect.Bottom
    End If
    '检测区域
    If dstRect.Right <= dstRect.Left Or dstRect.Bottom <= dstRect.Top Then Exit Function
    If srcRect.Right <= srcRect.Left Or srcRect.Bottom <= srcRect.Top Then Exit Function
    If m_EnableRecordRect Then m_DirtyRect = dstRect                '记录脏区域
    ClipperRect = True
End Function

'----------------------------- AlphaBlt -------------------------------
'处理透明显示,在默认alph=128时,速度将有所提高,单不能处理透明色
'SrcSurface  待显示的源表面
'x,y         显示的位置(相对关键点)
'alph        透明度(0-255)
'frame       SrcSurface中的待显示帧
'----------------------------------------------------------------------
Public Sub AlphaRectBlt(srcSurface As CSurface, ByVal x As Long, ByVal y As Long, srcRect As RECT, ByVal alph As Byte)
    Dim swidth As Long, sheight As Long, dstRect As RECT
    dstRect.Left = x
    dstRect.Top = y
    dstRect.Right = x + srcRect.Right - srcRect.Left
    dstRect.Bottom = y + srcRect.Bottom - srcRect.Top
    If Not ClipperRect(srcRect, dstRect, 1) Then Exit Sub
    On Error GoTo errhand:
    swidth = srcRect.Right - srcRect.Left
    sheight = srcRect.Bottom - srcRect.Top
    swidth = swidth And &HFFFC       '注意字节为4的倍数
    If swidth < 4 Then Exit Sub
    '注意在 ALHPA_SPEED_MODE 下可以不锁定直接操作
    Me.BeginLock
    srcSurface.BeginLock
    If alph = 128 Then
        '半透明
        halfablend_565_mmx ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, _
            ByVal srcSurface.DataAddress, srcRect.Left, srcRect.Top, _
            srcSurface.DataPitch, swidth, sheight
    Else
        ablend_565_mmx alph, ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, _
            ByVal srcSurface.DataAddress, srcRect.Left, srcRect.Top, srcSurface.DataPitch, _
            swidth, sheight, SYS_KEYCOLOR
    End If
    srcSurface.EndLock
    Me.EndLock

    
    Exit Sub
    
errhand:
    g_Err_Description = "AlphaBlt 失败!"
End Sub

Public Sub AlphaBlt(srcSurface As CSurface, ByVal x As Long, ByVal y As Long, Optional ByVal alph As Byte = 128, Optional frame As Long = 0)
    Dim srcRect As RECT, dstRect As RECT
    Dim swidth As Long, sheight As Long
    On Error GoTo errhand:
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, 1) Then Exit Sub
    swidth = srcRect.Right - srcRect.Left
    sheight = srcRect.Bottom - srcRect.Top
    swidth = swidth And &HFFFC       '注意字节为4的倍数
    If swidth < 4 Then Exit Sub
    '注意在 ALHPA_SPEED_MODE 下可以不锁定直接操作
    Me.BeginLock
    srcSurface.BeginLock
    If alph = 128 Then
        '半透明
        halfablend_565_mmx ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, _
            ByVal srcSurface.DataAddress, srcRect.Left, srcRect.Top, _
            srcSurface.DataPitch, swidth, sheight
    Else
        ablend_565_mmx alph, ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, _
            ByVal srcSurface.DataAddress, srcRect.Left, srcRect.Top, srcSurface.DataPitch, _
            swidth, sheight, SYS_KEYCOLOR
    End If
    srcSurface.EndLock
    Me.EndLock

    Exit Sub
    
errhand:
    g_Err_Description = "AlphaBlt 失败!"
    Debug.Print g_Err_Description
End Sub

Public Sub AlphaBltEx(srcSurface As CSurface, _
                    ByVal x As Long, _
                    ByVal y As Long, _
                    Optional ByVal alph As Byte = 128, _
                    Optional frame As Long = 0, _
                    Optional k As Single = 1)
    Dim srcRect As RECT, dstRect As RECT
    Dim swidth As Long, sheight As Long
    On Error GoTo errhand:
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, k) Then Exit Sub
    swidth = srcRect.Right - srcRect.Left
    sheight = srcRect.Bottom - srcRect.Top
    swidth = swidth And &HFFFC       '注意字节为4的倍数
    If swidth < 4 Then Exit Sub
    '注意在 ALHPA_SPEED_MODE 下可以不锁定直接操作
    Me.BeginLock
    srcSurface.BeginLock
    'DLL里面操作Alpha反了
    bltzoom_ablend_565_mmx 255 - alph, ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, _
            dstRect.Right - dstRect.Left, dstRect.Bottom - dstRect.Top, ByVal srcSurface.DataAddress, _
            srcRect.Left, srcRect.Top, srcSurface.DataPitch, _
            swidth, sheight, SYS_KEYCOLOR

    srcSurface.EndLock
    Me.EndLock

    Exit Sub
    
errhand:
    g_Err_Description = "AlphaBltEx 失败!"
    Debug.Print g_Err_Description
End Sub
'------------------------------------ Additive -----------------------------
'实现表面的光线叠加效果
'
'
'---------------------------------------------------------------------------
Public Sub Additive(srcSurface As CSurface, _
                    ByVal x As Long, _
                    ByVal y As Long, _
                    ByVal alph As Byte, _
                    Optional frame As Long = 0)
    Dim srcRect As RECT, dstRect As RECT
    Dim swidth As Long, sheight As Long
    On Error GoTo errhand:
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, 1) Then Exit Sub
    swidth = srcRect.Right - srcRect.Left
    sheight = srcRect.Bottom - srcRect.Top
    swidth = swidth And &HFFFC       '注意字节为4的倍数
    If swidth < 4 Then Exit Sub
    
    Me.BeginLock
    srcSurface.BeginLock
    additive_565_mmx alph, ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, _
        ByVal srcSurface.DataAddress, srcRect.Left, srcRect.Top, srcSurface.DataPitch, _
        swidth, sheight, SYS_KEYCOLOR

    srcSurface.EndLock
    Me.EndLock
    Exit Sub
    
errhand:
    Debug.Print Err.Description
End Sub

Public Sub AdditiveEx(srcSurface As CSurface, _
                    ByVal x As Long, _
                    ByVal y As Long, _
                    ByVal alph As Byte, _
                    Optional frame As Long = 0, _
                    Optional k As Single = 1)
    Dim srcRect As RECT, dstRect As RECT
    Dim swidth As Long, sheight As Long
    On Error GoTo errhand:
    
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, k) Then Exit Sub
    swidth = srcRect.Right - srcRect.Left
    sheight = srcRect.Bottom - srcRect.Top
    swidth = swidth And &HFFFC       '注意字节为4的倍数
    If swidth < 4 Then Exit Sub
    
    Me.BeginLock
    srcSurface.BeginLock
    
    bltzoom_additive_565_mmx alph, ByVal m_lpData, dstRect.Left, dstRect.Top, Me.DataPitch, _
            dstRect.Right - dstRect.Left, dstRect.Bottom - dstRect.Top, ByVal srcSurface.DataAddress, _
            srcRect.Left, srcRect.Top, srcSurface.DataPitch, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top, SYS_KEYCOLOR

    srcSurface.EndLock
    Me.EndLock
    Exit Sub
    
errhand:
    Debug.Print Err.Description
End Sub

Public Sub FastAdditive(srcSurface As CSurface, _
                    ByVal x As Long, _
                    ByVal y As Long, _
                    Optional frame As Long = 0)
    Dim srcRect As RECT, dstRect As RECT
    Dim swidth As Long, sheight As Long
    On Error GoTo errhand:
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, 1) Then Exit Sub
    
    swidth = srcRect.Right - srcRect.Left
    sheight = srcRect.Bottom - srcRect.Top
    swidth = swidth And &HFFFC       '注意字节为4的倍数
    If swidth < 4 Then Exit Sub
    
    Me.BeginLock
    srcSurface.BeginLock
    
    fast_additive_565_mmx ByVal srcSurface.DataAddress, _
            srcRect.Left, srcRect.Top, srcSurface.DataPitch, _
            ByVal m_lpData, dstRect.Left, dstRect.Top, Me.DataPitch, _
            swidth, sheight

    srcSurface.EndLock
    Me.EndLock
    Exit Sub
    
errhand:
    Debug.Print Err.Description
End Sub

Public Sub AddColorEx(srcSurface As CSurface, x As Long, y As Long, ByVal color As Long, Optional frame As Long = 0)
    Dim srcRect As RECT, dstRect As RECT
    Dim swidth As Long, sheight As Long
    On Error GoTo errhand:
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, 1) Then Exit Sub
    swidth = srcRect.Right - srcRect.Left
    sheight = srcRect.Bottom - srcRect.Top
    swidth = swidth And &HFFFC       '注意字节为4的倍数
    If swidth < 4 Then Exit Sub
    
    Me.BeginLock
    srcSurface.BeginLock
    addlightex_565_mmx ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, _
        ByVal srcSurface.DataAddress, srcRect.Left, srcRect.Top, srcSurface.DataPitch, _
        swidth, sheight, SYS_KEYCOLOR, color

    srcSurface.EndLock
    Me.EndLock
    Exit Sub
    
errhand:
    Debug.Print Err.Description
End Sub

'--------------------------------AddColor---------------------------------------------
'图象的快速加色运算
'-------------------------------------------------------------------------------------
Public Sub AddColor(ByVal x As Long, ByVal y As Long, ByVal W As Long, ByVal H As Long, ByVal color As Long, Optional frame As Long = 0)
    Dim dstRect As RECT, swidth As Long, sheight As Long
    Dim ddsd As DDSURFACEDESC2
    On Error GoTo errhand:
    dstRect.Left = x
    dstRect.Top = y
    dstRect.Right = x + W
    dstRect.Bottom = y + H
    If dstRect.Left < 0 Then dstRect.Left = 0
    If dstRect.Top < 0 Then dstRect.Top = 0
    If dstRect.Right > m_ClipperRect.Right Then dstRect.Right = m_ClipperRect.Right
    If dstRect.Bottom > m_ClipperRect.Bottom Then dstRect.Bottom = m_ClipperRect.Bottom
    swidth = dstRect.Right - dstRect.Left
    sheight = dstRect.Bottom - dstRect.Top
    swidth = swidth And &HFFFC

    If swidth < 4 Then Exit Sub
    Me.BeginLock
    
    addcolor_565_mmx ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, _
        swidth, sheight, color
    Me.EndLock
    Exit Sub
errhand:
    Debug.Print Err.Description
    'colorblend_565_mmx
End Sub

'--------------------------------LightAlphaBlt--------------------------------------
'常用来表现发光体的透明效果
'-----------------------------------------------------------------------------------
Public Sub LightAlphaBlt(srcSurface As CSurface, x As Long, y As Long, Optional frame As Long = 0)
    Dim srcRect As RECT, dstRect As RECT
    Dim swidth As Long, sheight As Long
    On Error GoTo errhand:
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, 1) Then Exit Sub
    swidth = srcRect.Right - srcRect.Left
    sheight = srcRect.Bottom - srcRect.Top

    Me.BeginLock
    srcSurface.BeginLock
    swidth = swidth And &HFFFC
    If swidth < 4 Then Exit Sub
    
    light_565_mmx ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, _
        ByVal srcSurface.DataAddress, srcRect.Left, srcRect.Top, srcSurface.DataPitch, _
        swidth, sheight, SYS_KEYCOLOR

    srcSurface.EndLock
    Me.EndLock
    Exit Sub
errhand:
    Debug.Print Err.Description
End Sub

Public Sub Subtract(srcSurface As CSurface, x As Long, y As Long, ByVal alph As Byte, Optional frame As Long = 0)
    Dim srcRect As RECT, dstRect As RECT
    Dim swidth As Long, sheight As Long
    On Error GoTo errhand:
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, 1) Then Exit Sub
    swidth = srcRect.Right - srcRect.Left
    sheight = srcRect.Bottom - srcRect.Top
    swidth = swidth And &HFFFC       '注意字节为4的倍数
    If swidth < 4 Then Exit Sub
    
    Me.BeginLock
    srcSurface.BeginLock
    subitive_565_mmx alph, ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, _
        ByVal srcSurface.DataAddress, srcRect.Left, srcRect.Top, srcSurface.DataPitch, _
        swidth, sheight, SYS_KEYCOLOR

    srcSurface.EndLock
    Me.EndLock
    Exit Sub
    
errhand:
    Debug.Print Err.Description
End Sub

Public Sub MixColor(srcSurface As CSurface, ByVal x As Long, ByVal y As Long, ByVal color As Long, Optional frame As Long = 0)
    '处理图象混色加速,使用Alpha=0.5
    Dim srcRect As RECT, dstRect As RECT
    Dim swidth As Long, sheight As Long
    On Error GoTo errhand:
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, 1) Then Exit Sub
    swidth = srcRect.Right - srcRect.Left
    sheight = srcRect.Bottom - srcRect.Top
    
    Me.BeginLock
    srcSurface.BeginLock
    colorblend_565 255, ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, _
        ByVal srcSurface.DataAddress, srcRect.Left, srcRect.Top, srcSurface.DataPitch, _
        swidth, sheight, SYS_KEYCOLOR, color
    srcSurface.EndLock
    Me.EndLock
    Exit Sub
errhand:
    Debug.Print Err.Description
End Sub

Public Sub GreenMask(srcSurface As CSurface, ByVal mask As Long, ByVal x As Long, ByVal y As Long, Optional frame As Long = 0)
    '处理图象混色加速,使用Alpha=0.5
    Dim srcRect As RECT, dstRect As RECT
    Dim swidth As Long, sheight As Long
    On Error GoTo errhand:
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, 1) Then Exit Sub
    
    swidth = srcRect.Right - srcRect.Left
    sheight = srcRect.Bottom - srcRect.Top
    
    Me.BeginLock
    srcSurface.BeginLock
    swidth = swidth And &HFFFC
    If swidth < 4 Then Exit Sub
    
    mask_565_mmx ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, _
        ByVal srcSurface.DataAddress, srcRect.Left, srcRect.Top, srcSurface.DataPitch, _
        swidth, sheight, mask, SYS_KEYCOLOR
        
    srcSurface.EndLock
    Me.EndLock
    Exit Sub
errhand:
    Debug.Print Err.Description
End Sub

Public Sub DrawAlphaRect(ByVal x As Long, ByVal y As Long, ByVal W As Long, ByVal H As Long, ByVal alph As Byte)
    '显示一个阴影矩形
    Dim dstRect As RECT, swidth As Long, sheight As Long
    Dim ddsd As DDSURFACEDESC2
    On Error GoTo errhand:
    dstRect.Left = x
    dstRect.Top = y
    dstRect.Right = x + W
    dstRect.Bottom = y + H
    If dstRect.Left < 0 Then dstRect.Left = 0
    If dstRect.Top < 0 Then dstRect.Top = 0
    If dstRect.Right > m_ClipperRect.Right Then dstRect.Right = m_ClipperRect.Right
    If dstRect.Bottom > m_ClipperRect.Bottom Then dstRect.Bottom = m_ClipperRect.Bottom
    swidth = dstRect.Right - dstRect.Left
    sheight = dstRect.Bottom - dstRect.Top
    swidth = swidth And &HFFFC
    If g_Mode = ALHPA_SPEED_MODE Then
        'If m_lpdata = 0 Then Exit Sub
        alpharect_565_mmx alph, ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, swidth, sheight
    Else
        Me.BeginLock
        alpharect_565_mmx alph, ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, swidth, sheight
        Me.EndLock
    End If
    Exit Sub
    
errhand:
    Debug.Print Err.Description
End Sub

'----------------------------- Clear --------------------------------
'   用黑色填充表面
'--------------------------------------------------------------------
Public Sub Clear()
    
    If g_Windowed Or g_Mode = BLT_SPEED_MODE Then
        'StickStart
        m_Surface.BltColorFill m_ClipperRect, 0
        'Debug.Print StickEnd; "BltColorFill"
    Else
        'StickStart
        Qmemset ByVal m_lpData, 0, m_width * CLng(m_height) / 4
        'Debug.Print StickEnd; "Qmemset"
    End If
End Sub

'---------------------------- FastBlt/Blt ---------------------------------
'实现表面之间的数据块传送
'SrcSurface         源表面
'x,y                显示的位置(相对关键点)
'Frame              SrcSurface中的待显示帧
'k                  缩放系数
'----------------------------------------------------------------------
Public Sub FastBltEx(srcSurface As CSurface, x As Long, y As Long, Optional frame As Long = 0)
    Dim srcRect As RECT, dstRect As RECT
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, 1) Then Exit Sub
    
    If g_Mode = ALHPA_SPEED_MODE Then
        bltfast_mmx ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, dstRect.Right - dstRect.Left, _
            dstRect.Bottom - dstRect.Top, ByVal srcSurface.DataAddress, srcRect.Left, srcRect.Top, _
            srcSurface.DataPitch, SYS_KEYCOLOR
    Else
        m_Surface.bltfast dstRect.Left, dstRect.Top, srcSurface.DD_Surface, srcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End If
End Sub

Public Sub FastBlt(srcSurface As CSurface, x As Long, y As Long, Optional frame As Long = 0)
    Dim srcRect As RECT, dstRect As RECT
    '不处理缩放和透明，直接使用DX7

    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, 1) Then Exit Sub
    'm_Surface.bltfast dstRect.Left, dstRect.Top, srcsurface.DD_Surface, srcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    
    If g_Mode = ALHPA_SPEED_MODE Then
        bltfast ByVal m_lpData, dstRect.Left, dstRect.Top, m_lPitch, dstRect.Right - dstRect.Left, _
            dstRect.Bottom - dstRect.Top, ByVal srcSurface.DataAddress, srcRect.Left, srcRect.Top, _
            srcSurface.DataPitch
    Else
        m_Surface.bltfast dstRect.Left, dstRect.Top, srcSurface.DD_Surface, srcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    End If
End Sub

'----------------------RLE透明压缩包的支持--------------------------------------
'如果需要使用ＲＬＥ支持，请将该代码启用，并且添加相应的类模块
'
'Public Sub RleBit(srcSurface As CRleSurface, x As Long, y As Long)
    'Rle 解压
'    Me.BeginLock
'    rle_blt ByVal m_lpData, Me.DataPitch, m_height, x, y, _
'            ByVal srcSurface.GetlpRleDatas, srcSurface.GetRlePointsNum
'    Me.EndLock
'End Sub

Public Sub Blt(srcSurface As CSurface, ByVal x As Long, ByVal y As Long, Optional ByVal frame As Long = 0, _
    Optional ByVal k As Single = 1)
    Dim srcRect As RECT, dstRect As RECT
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, k) Then Exit Sub
    Me.BeginLock
    srcSurface.BeginLock
    'bltfast ByVal m_lpData, dstRect.Left, dstRect.Top, Me.DataPitch, _
            dstRect.Right - dstRect.Left, dstRect.Bottom - dstRect.Top, ByVal SrcSurface.DataAddress, _
            srcRect.Left, srcRect.Top, SrcSurface.DataPitch
    'bltfast_mmx ByVal m_lpData, dstRect.Left, dstRect.Top, Me.DataPitch, _
            dstRect.Right - dstRect.Left, dstRect.Bottom - dstRect.Top, ByVal SrcSurface.DataAddress, _
            srcRect.Left, srcRect.Top, SrcSurface.DataPitch, SYS_KEYCOLOR
    'scanx_565 ByVal SrcSurface.DataAddress, srcRect.Left, srcRect.Top, SrcSurface.DataPitch, _
             ByVal m_lpData, dstRect.Left, dstRect.Top, Me.DataPitch, _
            srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top, _
            &H1F, SYS_KEYCOLOR
    'scan_linexy ByVal m_lpData, dstRect.Left, dstRect.Top, Me.DataPitch, _
            ByVal SrcSurface.DataAddress, srcRect.Left, srcRect.Top, SrcSurface.DataPitch, _
            srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top, _
            &H1F, SYS_KEYCOLOR
    bltzoom_565_mmx ByVal m_lpData, dstRect.Left, dstRect.Top, Me.DataPitch, _
            dstRect.Right - dstRect.Left, dstRect.Bottom - dstRect.Top, ByVal srcSurface.DataAddress, _
            srcRect.Left, srcRect.Top, srcSurface.DataPitch, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top, SYS_KEYCOLOR
    srcSurface.EndLock
    Me.EndLock
    
    'm_Surface.Blt dstRect, SrcSurface.DD_Surface, srcRect, DDBLT_KEYSRC Or DDBLT_WAIT
End Sub

Public Sub BltWithEdgeline(srcSurface As CSurface, ByVal x As Long, ByVal y As Long, ByVal color As Long, Optional frame As Long = 0)
    Dim srcRect As RECT, dstRect As RECT

    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, 1) Then Exit Sub
    Me.BeginLock
    srcSurface.BeginLock
    
    scan_linexy ByVal m_lpData, dstRect.Left, dstRect.Top, Me.DataPitch, _
            ByVal srcSurface.DataAddress, srcRect.Left, srcRect.Top, srcSurface.DataPitch, _
            srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top, _
            color, SYS_KEYCOLOR
    
    srcSurface.EndLock
    Me.EndLock
End Sub

'--------------------------------镜像Blt---------------------------------
'MirrorLeftRightBlt： 左右翻转
'MirrorUpDownBlt：    上下翻转（请在Mode=Alpha_Blt_Speed下使用）
'------------------------------------------------------------------------
Public Sub MirrorLeftRightBlt(srcSurface As CSurface, x As Long, y As Long, Optional frame As Long = 0, _
    Optional k As Single = 1)
    Dim srcRect As RECT, dstRect As RECT, ddfx As DDBLTFX
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, k) Then Exit Sub
    ddfx.lDDFX = DDBLTFX_MIRRORLEFTRIGHT
    m_Surface.BltFx dstRect, srcSurface.DD_Surface, srcRect, DDBLT_KEYSRC Or DDBLT_WAIT Or DDBLT_DDFX, ddfx
End Sub

Public Sub MirrorUpDownBlt(srcSurface As CSurface, x As Long, y As Long, Optional frame As Long = 0, _
    Optional k As Single = 1)
    Dim srcRect As RECT, dstRect As RECT, ddfx As DDBLTFX
    If Not ClippeSurface(srcSurface, x, y, srcRect, dstRect, frame, k) Then Exit Sub
    ddfx.lDDFX = DDBLTFX_MIRRORUPDOWN
    m_Surface.BltFx dstRect, srcSurface.DD_Surface, srcRect, DDBLT_KEYSRC Or DDBLT_WAIT Or DDBLT_DDFX, ddfx
End Sub

Public Sub RotateRect(srcSurface As CSurface, ByVal x As Long, ByVal y As Long, angle As Single, Optional frame As Long = 0)
    '注意中心点为原点，注意不要越界,小心使用
    Dim swidth As Long, sheight As Long
    Dim sx As Long, sy As Long, keyx As Long, keyy As Long, srcRect As RECT
    
    srcRect = srcSurface.FrameRect(frame)
    
    sx = srcRect.Left
    sy = srcRect.Top

    swidth = srcRect.Right - srcRect.Left
    sheight = srcRect.Bottom - srcRect.Top
    
    '防止越界
    Debug.Assert x >= 0 And y >= 0 And x + swidth < m_width And y + sheight < m_height
    
    Me.BeginLock
    srcSurface.BeginLock
    rotate_tran ByVal m_lpData, Me.DataPitch, ByVal srcSurface.DataAddress, srcSurface.DataPitch, _
        x + swidth / 2, y + sheight / 2, swidth, sheight, _
        sx, sy, swidth, sheight, angle, SYS_KEYCOLOR
        '源中心为 sx+swidth/2,sy+swidth/2
    srcSurface.EndLock
    Me.EndLock
End Sub

'--------------------------------- RectBlt ---------------------------------
'直接在两表面之间传送指定区域数据
'SrcSurface        源表面
'dstRect           目标矩形区域
'srcRect           源矩形区域
'---------------------------------------------------------------------------
Public Sub RectBlt(srcSurface As CSurface, dstRect As RECT, srcRect As RECT)
    Dim swidth As Long, sheight As Long, dwidth As Long, dheight As Long
    Dim kX As Single, kY As Single, dx As Long, dy As Long
    'On Error Resume Next
    kX = (srcRect.Right - srcRect.Left) / (dstRect.Right - dstRect.Left)
    kY = (srcRect.Bottom - srcRect.Top) / (dstRect.Bottom - dstRect.Top)
    
    'Clipper m_ClipperRect
    If dstRect.Left < m_ClipperRect.Left Then dx = m_ClipperRect.Left - dstRect.Left: dstRect.Left = m_ClipperRect.Left
    If dstRect.Top < m_ClipperRect.Top Then dy = m_ClipperRect.Top - dstRect.Top: dstRect.Top = m_ClipperRect.Top
    If dstRect.Right > m_ClipperRect.Right Then dstRect.Right = m_ClipperRect.Right
    If dstRect.Bottom > m_ClipperRect.Bottom Then dstRect.Bottom = m_ClipperRect.Bottom
    
    dwidth = dstRect.Right - dstRect.Left
    dheight = dstRect.Bottom - dstRect.Top
    If dwidth = 0 Or dheight = 0 Then Exit Sub
    
    With srcRect
        .Left = .Left + dx * kX
        .Top = .Top + dy * kY
        .Right = .Left + dwidth * kX
        .Bottom = .Top + dheight * kY
    End With

    If srcRect.Bottom > srcSurface.Height Then srcRect.Bottom = srcSurface.Height
    If srcRect.Right > srcSurface.Width Then srcRect.Right = srcSurface.Width
    
    m_Surface.Blt dstRect, srcSurface.DD_Surface, srcRect, DDBLT_KEYSRC Or DDBLT_WAIT
End Sub

Public Sub RectBltEx(srcSurface As CSurface, ByVal x As Long, ByVal y As Long, _
    ByVal W As Long, ByVal H As Long, ByVal sx As Long, ByVal sy As Long, _
    Optional ByVal sw As Long, Optional ByVal sh As Long)
    Dim dstRect As RECT, srcRect As RECT
    Dim dx As Long, dy As Long, kX As Single, kY As Single
    On Error GoTo errhand
    
    kX = IIf(sw = 0, 1, sw / W): kY = IIf(sh = 0, 1, sh / H)
    
    dstRect.Left = x
    dstRect.Top = y
    dstRect.Right = x + W
    dstRect.Bottom = y + H
    
    srcRect.Left = sx
    srcRect.Top = sy
    srcRect.Right = sx + sw
    srcRect.Bottom = sy + sh
    'Clipper m_ClipperRect
    If dstRect.Left < m_ClipperRect.Left Then dx = m_ClipperRect.Left - dstRect.Left: dstRect.Left = m_ClipperRect.Left
    If dstRect.Top < m_ClipperRect.Top Then dy = m_ClipperRect.Top - dstRect.Top: dstRect.Top = m_ClipperRect.Top
    If dstRect.Right > m_ClipperRect.Right Then dstRect.Right = m_ClipperRect.Right
    If dstRect.Bottom > m_ClipperRect.Bottom Then dstRect.Bottom = m_ClipperRect.Bottom
    
    W = dstRect.Right - dstRect.Left
    H = dstRect.Bottom - dstRect.Top
    If W = 0 Or H = 0 Then Exit Sub
    
    With srcRect
        .Left = .Left + dx * kX
        .Top = .Top + dy * kY
        .Right = .Left + W * kX
        .Bottom = .Top + H * kY
    End With

    If srcRect.Bottom > srcSurface.Height Then srcRect.Bottom = srcSurface.Height
    If srcRect.Right > srcSurface.Width Then srcRect.Right = srcSurface.Width
    
    m_Surface.Blt dstRect, srcSurface.DD_Surface, srcRect, DDBLT_KEYSRC Or DDBLT_WAIT
errhand:
    
End Sub

Public Function ReadFrames(ByVal FileName As String) As Boolean
    Dim keyx As Integer, keyy As Integer, FileNumber As Long
    Dim srcRect As RECT
    Dim sFile As String
    Dim FrameNum As Integer
    ReDim m_Frames(0)
    m_Frames(0).Area.Right = m_width
    m_Frames(0).Area.Bottom = m_height
    '查找对应的Dat文件若有读取--------------A Bug May Happened
    sFile = Mid(FileName, 1, Len(FileName) - 3) + "dat"
    If Dir(sFile) = "" Then
        ReadFrames = True: Exit Function
    End If
    
    On Error GoTo errhand
    FileNumber = FreeFile
    Open sFile For Binary As #FileNumber

    Get #FileNumber, , FrameNum
    ReDim m_Frames(FrameNum - 1)
    Get #FileNumber, , m_Frames

    Close #FileNumber
    ReadFrames = True
    Exit Function
    
errhand:
    Close #1
    g_Err_Description = "读取Dat文件出错！"
End Function

Public Sub AddFrame(srcRect As RECT, keyx As Long, keyy As Long)
    m_Frames(UBound(m_Frames())).Area = srcRect
    m_Frames(UBound(m_Frames())).KeyPointX = keyx
    m_Frames(UBound(m_Frames())).KeyPointY = keyy
    ReDim Preserve m_Frames(UBound(m_Frames()) + 1)
End Sub

Public Function FrameRect(ByVal FrameCount As Long) As RECT
    If FrameCount > UBound(m_Frames()) Or FrameCount < 0 Then Exit Function
    FrameRect = m_Frames(FrameCount).Area
End Function

Public Sub GetFrameKeyPoint(ByVal FrameCount As Long, x As Long, y As Long)
    If FrameCount > UBound(m_Frames()) Or FrameCount < 0 Then Exit Sub
    x = m_Frames(FrameCount).KeyPointX
    y = m_Frames(FrameCount).KeyPointY
End Sub

Public Sub SetForeColor(color As Long)
    m_ForeColor = color
    m_Surface.SetForeColor color
End Sub

Public Sub Render(ByVal dsthdc As Long, Optional frame As Long = 0)
    Dim thdc As Long, srcRect As RECT
    thdc = m_Surface.GetDC
    srcRect = Me.FrameRect(frame)
    BitBlt dsthdc, 0, 0, srcRect.Right - srcRect.Left, srcRect.Bottom - srcRect.Top, _
        thdc, srcRect.Left, srcRect.Top, vbSrcCopy
    DD_Surface.ReleaseDC thdc
End Sub

Public Sub RenderTo(ByVal dsthdc As Long)
    Dim thdc As Long
    thdc = m_Surface.GetDC
    BitBlt dsthdc, 0, 0, m_width, m_height, thdc, 0, 0, vbSrcCopy
    DD_Surface.ReleaseDC thdc
End Sub

