Attribute VB_Name = "mdlDirectX"
Option Explicit

'//Adapter Identifier
Private EnumIdent8 As D3DADAPTER_IDENTIFIER8
    
'//Enumeration for the device capabilities
Private EnumD3DCaps8 As D3DCAPS8

'//Enumeration for the display modes
Private EnumD3DMode As D3DDISPLAYMODE

'//Used to store the device type, for creation of the device
Private EnumDeviceType As CONST_D3DDEVTYPE

'//Window format
Private EnumModeFormat As CONST_D3DFORMAT

Public EnumD3DPresent As D3DPRESENT_PARAMETERS      '//Another enumeration for Direct3D, containing information about the Video card


Private lngVertexBehavior As Long

Private lngWindowWidth As Long, lngWindowHeight As Long
Private lngWindowHalfWidth As Long, lngWindowHalfHeight As Long

Public lngSetupWidth As Long, lngSetupHeight As Long '//As configured

Public blnIsWindowed As Boolean

Public Property Let WindowWidth(lngWidth As Long)
    lngWindowWidth = lngWidth
    
    '//Also calculate half width
    lngWindowHalfWidth = lngWidth \ 2
End Property

Public Property Get WindowWidth() As Long
    WindowWidth = lngWindowWidth
End Property
'------------
Public Property Let WindowHeight(lngHeight As Long)
    lngWindowHeight = lngHeight
    '//Also calculate half height
    lngWindowHalfHeight = lngHeight \ 2
End Property

Public Property Get WindowHeight() As Long
    WindowHeight = lngWindowHeight
End Property

'------------
Public Property Let WindowHalfWidth(lngWidth As Long)
    lngWindowHalfWidth = lngWidth
End Property

Public Property Get WindowHalfWidth() As Long
    WindowHalfWidth = lngWindowHalfWidth
End Property
'------------
Public Property Let WindowHalfHeight(lngHeight As Long)
    lngWindowHalfHeight = lngHeight
End Property

Public Property Get WindowHalfHeight() As Long
    WindowHalfHeight = lngWindowHalfHeight
End Property

Private Sub Class_Terminate()
    '//Automatically unload DirectX objects
    UnloadDirectXObjects
End Sub

Public Sub InitDirectXObjects()
    
    '//Create new DirectX8 object
    '//Will be used through the entire project
    Set ObjDX = New DirectX8
    
    '//Create D3D Object
    Set objD3D = ObjDX.Direct3DCreate
    
    '//Create D3DX helper object
    Set ObjD3DX = New D3DX8

End Sub

Private Sub UnloadDirectXObjects()
    On Local Error Resume Next
    
    Set ObjD3DX = Nothing
    Set objD3D = Nothing
    Set ObjDX = Nothing
End Sub

Public Sub InitDirectGraphics(lngFormHandle As Long, Optional blnWindowed As Boolean = False)
    Dim lngFlags As Long
    Dim intDeviceID As Integer
    Dim lngAdapter As Long
    
    '//Get adapter identifier
    objD3D.GetAdapterIdentifier lngAdapter, 0, EnumIdent8
    
    EnumDeviceType = D3DDEVTYPE_HAL
    Call objD3D.GetDeviceCaps(lngAdapter, EnumDeviceType, EnumD3DCaps8)
        If Err.Number Then
            Err.Clear
            '//Error received, try Reference rasterizer as device
            EnumDeviceType = D3DDEVTYPE_REF
            Call objD3D.GetDeviceCaps(lngAdapter, EnumDeviceType, EnumD3DCaps8)
                If Err.Number Then
                    Err.Clear
                    '//Reference rasterizer doesn't work either.
                    MsgBox "No HAL or REF device found. This game cannot run.", vbOKOnly + vbCritical
                    End
                End If
        End If

    
    Call objD3D.GetAdapterDisplayMode(lngAdapter, EnumD3DMode)


    ResetWindowState blnWindowed
    
    '//Store
    blnIsWindowed = blnWindowed
    '//Only use software vertex processing (no T&L)
    lngFlags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
    lngVertexBehavior = lngFlags

    Set ObjDev = objD3D.CreateDevice(lngAdapter, EnumDeviceType, lngFormHandle, lngFlags, EnumD3DPresent)
        '//Couldn't create the device
        If Err.Number Then
            Err.Clear
            '//Try again using Reference rasterizer
            Set ObjDev = objD3D.CreateDevice(lngAdapter, D3DDEVTYPE_REF, lngFormHandle, lngFlags, EnumD3DPresent)
            If Err.Number Then
                '//No luck either. Quit
                MsgBox "No HAL or REF device found. This game cannot run.", vbOKOnly + vbCritical
                End
            End If
        End If
    
        
    
'    CheckCaps
    
    '//Set some texture/render states
    ResetStates

End Sub

Public Sub ResetStates()
    With ObjDev
        Call .SetVertexShader(D3DFVF_TLVERTEX)
        Call .SetRenderState(D3DRS_LIGHTING, False)
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        Call .SetRenderState(D3DRS_ZENABLE, False)
        Call .SetRenderState(D3DRS_ZWRITEENABLE, False)
        Call .SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
        Call .SetRenderState(D3DRS_SHADEMODE, D3DSHADE_FLAT)
        Call .SetRenderState(D3DRS_FILLMODE, CONST_D3DFILLMODE.D3DFILL_SOLID)
        
        .SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
        .SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
        .SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        Call .SetTextureStageState(0, D3DTSS_MINFILTER, D3DTEXF_LINEAR)
        Call .SetTextureStageState(0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR)

    End With
End Sub

'Public Sub CheckCaps()
'    Dim myDeviceCaps As D3DCAPS8
'    Dim strString(1) As String, strResult(1) As String
'
'    strString(0) = "Device supports: SRCBLEND = "
'    strString(1) = "Device supports: DESTBLEND = "
'    strResult(0) = "-> NO!"
'    strResult(1) = "-> OK!"
'    '//Get Device caps
'    ObjDev.GetDeviceCaps myDeviceCaps
'
'    With myDeviceCaps
'        '//Check if device supports one-one blending
'        If (0 = (myDeviceCaps.SrcBlendCaps And D3DPBLENDCAPS_ONE) = D3DPBLENDCAPS_ONE) Then
'            mdlLog.AddToLog strString(0) & "D3DPBLENDCAPS_ONE" & strResult(0)
'        Else
'            mdlLog.AddToLog strString(0) & "D3DPBLENDCAPS_ONE" & strResult(1)
'        End If
'
'        If (0 = (myDeviceCaps.DestBlendCaps And D3DPBLENDCAPS_ONE) = D3DPBLENDCAPS_ONE) Then
'            mdlLog.AddToLog strString(1) & "D3DPBLENDCAPS_ONE" & strResult(0)
'        Else
'            mdlLog.AddToLog strString(1) & "D3DPBLENDCAPS_ONE" & strResult(1)
'        End If
'
'        If (0 = (myDeviceCaps.SrcBlendCaps And D3DPBLENDCAPS_SRCALPHA) = D3DPBLENDCAPS_SRCALPHA) Then
'            mdlLog.AddToLog strString(0) & "D3DPBLENDCAPS_SRCALPHA" & strResult(0)
'        Else
'            mdlLog.AddToLog strString(0) & "D3DPBLENDCAPS_SRCALPHA" & strResult(1)
'        End If
'
'        If (0 = (myDeviceCaps.DestBlendCaps And D3DPBLENDCAPS_INVSRCALPHA) = D3DPBLENDCAPS_INVSRCALPHA) Then
'            mdlLog.AddToLog strString(1) & "D3DPBLENDCAPS_INVSRCALPHA" & strResult(0)
'        Else
'            mdlLog.AddToLog strString(1) & "D3DPBLENDCAPS_INVSRCALPHA" & strResult(1)
'        End If
'
'        mdlLog.AddToLog "Device: Maximum texture repeat: " & myDeviceCaps.MaxTextureRepeat
'
'        If (0 = (myDeviceCaps.TextureCaps And D3DPTEXTURECAPS_TEXREPEATNOTSCALEDBYSIZE) = D3DPTEXTURECAPS_TEXREPEATNOTSCALEDBYSIZE) Then
'            mdlLog.AddToLog "Device: D3DPTEXTURECAPS_TEXREPEATNOTSCALEDBYSIZE is not set"
'            mdlLog.AddToLog "Device: Max texture 32x32 repeat is therefore: " & myDeviceCaps.MaxTextureRepeat \ 32
'            mdlLog.AddToLog "Device: Range [-" & (myDeviceCaps.MaxTextureRepeat \ 32) \ 2 & "," & (myDeviceCaps.MaxTextureRepeat \ 32) \ 2 & "]"
'        Else
'            mdlLog.AddToLog "Device: D3DPTEXTURECAPS_TEXREPEATNOTSCALEDBYSIZE is set"
'            mdlLog.AddToLog "Device: Max texture 32x32 repeat is therefore: " & myDeviceCaps.MaxTextureRepeat
'            mdlLog.AddToLog "Device: Texture Range [-" & (myDeviceCaps.MaxTextureRepeat) \ 2 & "," & (myDeviceCaps.MaxTextureRepeat) \ 2 & "]"
'        End If
'        mdlLog.AddToLog "This game uses: [-16,16] max"
'   End With
'End Sub

Public Sub ResetWindowState(blnWindowed As Boolean)

 '   If blnWindowed = False Then
        With EnumD3DPresent
            .BackBufferFormat = EnumD3DMode.Format
            .SwapEffect = D3DSWAPEFFECT_FLIP
            .BackBufferCount = 1
            .BackBufferWidth = WindowWidth
            .BackBufferHeight = WindowHeight
            .hDeviceWindow = frmMain.hwnd
            .Windowed = 0

            .FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_DEFAULT
            'D3DPRESENT_INTERVAL_IMMEDIATE

            '.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
        End With
'    Else


'    Dim intRealWidth As Integer, intRealHeight As Integer
'    intRealWidth = 800
'    intRealHeight = 600
'
'    intClientHeight = intRealHeight
'    intClientWidth = intRealWidth
'
'
'
'        With EnumD3DPresent
'            .BackBufferFormat = EnumD3DMode.Format
'            .SwapEffect = D3DSWAPEFFECT_DISCARD
'            .BackBufferCount = 1
'            .BackBufferWidth = intClientWidth
'            .BackBufferHeight = intClientHeight
'            .EnableAutoDepthStencil = 1
'            .AutoDepthStencilFormat = D3DFMT_D16
'            .Windowed = 1
'            .hDeviceWindow = frmMain.hwnd
'            .FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_DEFAULT
'        End With
'
'        With frmMain
'            .Show
'            .Caption = AppTitle
'
'            .Width = .ScaleX(intRealWidth, vbPixels, vbTwips)
'            .Height = .ScaleY(intRealHeight, vbPixels, vbTwips)
'
'            intRealWidth = .ScaleX(.Width, vbTwips, vbPixels)
'            intRealHeight = .ScaleY(.Height, vbTwips, vbPixels)
'
'            intRealWidth = intRealWidth + (intRealWidth - .ScaleWidth)
'            intRealHeight = intRealHeight + (intRealHeight - .ScaleHeight)
'
'            .Width = .ScaleX(intRealWidth, vbPixels, vbTwips)
'            .Height = .ScaleY(intRealHeight, vbPixels, vbTwips)
'
'            .Left = (intOriginalResWidthTwip / 2) - (.Width / 2)
'            .Top = (intOriginalResHeightTwip / 2) - (.Height / 2)
'        End With
'    End If


End Sub
'
'Public Sub WindowResized()
'    '//Reset the device to the new mode
'    Call mdlTools.DeviceReset
'End Sub
'
'Public Function GenerateResolutionsList() As String()
'
'End Function

