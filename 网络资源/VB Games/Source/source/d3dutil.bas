Attribute VB_Name = "D3DUtil"


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Copyright (C) 1999-2001 Microsoft Corporation.  All Rights Reserved.
'
' File:       D3DUtil.Bas
' Content:    VB D3DFramework utility module
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'DOC:  Use with
'DOC:        D3DAnimation.cls
'DOC:        D3DFrame.cls
'DOC:        D3DMesh.cls
'DOC:        D3DSelectDevice.frm (optional)
'DOC:
'DOC:  Short list of usefull functions
'DOC:        D3DUtil_Init                  first call to framework
'DOC:        D3DUtil_LoadFromFile          loads an x-file
'DOC:        D3DUtil_SetupDefaultScene     setup a camera lights and materials
'DOC:        D3DUtil_SetupCamera           point camera
'DOC:        D3DUtil_SetupMediaPath        set directory to load textures from
'DOC:        D3DUtil_PresentAll            show graphic on the screen
'DOC:        D3DUtil_ResizeWindowed        resize for windowed modes
'DOC:        D3DUtil_ResizeFullscreen      resize to fullscreen mode
'DOC:        D3DUtil_CreateTextureInPool   create a texture


Option Explicit


'DOC: DXLockArray8 & DXUnlockArray8
'DOC:
'DOC: These are Helper functions that allow textures, vertex buffers, and index buffers
'DOC: to look like VB arrays to the VB user.
'DOC: It is imperative that Lock be matched with unlock or undefined behaviour may result
'DOC: It is imperative that DXLockarray8 be matched with DXUnlockArray8 or undefined behaviour may result
'DOC:
'DOC: DXLockArray8
'DOC:       resource    - can be Direct3DTexture8,Direct3dVertexBuffer8, or a Direct3DIndexBuffer
'DOC:       addr        - is the number provide by IndexBuffer.Lock,Testure.Lock etc
'DOC:       arr()       - a VB array that can be used to shadow video memory
'DOC: DXUnlockArray8
'DOC:       resource    - the resource passed to DXLockArray8
'DOC:       arr()       - the VB array passed to DXLockArray8
'DOC:
'DOC: Example
'DOC:           dim m_vertBuff as Direct3DVertexBuffer  'we assume this has been created
'DOC:           dim m_vertCount as long                 'we assume this has been set
'DOC:
'DOC:           Dim addr As Long                        'will holds the address the D3D
'DOC:                                                   'managed memory
'DOC:           dim verts() as D3DVERTEX                'array that we want to point to
'DOC:                                                   'D3D managed memory
'DOC:
'DOC:           redim verts(m_vertCount)                'ensure the size is large
'DOC:                                                   'enough for the data and has
'DOC:                                                   'as many dimensions as needed
'DOC:                                                   '(1d for vertex buffer, 2d for
'DOC:                                                   'surfaces, 3d for volumes)
'DOC:                                                   'resize the array once and
'DOC:                                                   'reuse for frequent manipulation
'DOC:
'DOC:           m_vertBuff.Lock 0, Len(verts(0)) * m_vertCount, addr, 0
'DOC:
'DOC:           DXLockArray8 m_vertBuff, addr, verts
'DOC:
'DOC:           for i = 0 to m_vertCount-1
'DOC:               verts(i).x=i 'or what ever you want to dow with the data
'DOC:           next
'DOC:
'DOC:           DXUnlockArray8 m_vertBuff, verts
'DOC:
'DOC:           VB.Unlock
'
Public Declare Function DXLockArray8 Lib "dx8vb.dll" (ByVal resource As Direct3DResource8, ByVal addr As Long, arr() As Any) As Long
Public Declare Function DXUnlockArray8 Lib "dx8vb.dll" (ByVal resource As Direct3DResource8, arr() As Any) As Long



'DOC: Texture Load data applied to all textures
'DOC: can be accessed by g_TextureSampling variable
Private Type TextureParams
    enable As Boolean           'enable texture sampling
    
    width As Long               'default width of textures
    height As Long              'default height of textures
    miplevels As Long           'default number of miplevels
    mipfilter As Long           'default mipmap filter
    filter As Long              'default texture filter
    fmt As CONST_D3DFORMAT      'default texture format
    fmtTrans As CONST_D3DFORMAT 'default transparent format
    colorTrans As Long          'default transparent color
    
End Type


'DOC: Rotate key used in conjuction with the CD3DAnimation class
Public Type D3DROTATEKEY
    time As Long
    nFloats As Long
    quat As D3DQUATERNION
End Type

'DOC: Scale or Translate key used in conjuction with the CD3DAnimation class
Public Type D3DVECTORKEY
    time As Long
    nFloats As Long
    vec As D3DVECTOR
End Type

'DOC: Pick record using with CD3DPick class
Public Type D3D_PICK_RECORD
    hit As Long
    triFaceid As Long
    a       As Single
    b       As Single
    dist   As Single
End Type

'DOC: see D3DUtil_Timer
Public Enum TIMER_COMMAND
          TIMER_RESET = 1         '- to reset the timer
          TIMER_start = 2         '- to start the timer
          TIMER_STOP = 3          '- to stop (or pause) the timer
          TIMER_ADVANCE = 4       '- to advance the timer by 0.1 seconds
          TIMER_GETABSOLUTETIME = 5 '- to get the absolute system time
          TIMER_GETAPPTIME = 6      '- to get the current time
          TIMER_GETELLAPSEDTIME = 7 '- to get the ellapsed time
End Enum


'DOC: Info on a per texture basis
Private Type TexPoolEntry
    Name As String
    tex As Direct3DTexture8
    nextDelNode As Long
End Type



'------------------------------------------------------------------
'DOC: Usefull globals
'------------------------------------------------------------------


Public g_bDontDrawTextures As Boolean           'Debuging switches
Public g_bClipMesh As Boolean                   'Debuging switches
Public g_bLoadSkins  As Boolean                 'Debuging switches
Public g_bLoadNoAlpha As Boolean                'Debuging switches

                                                'view frustrum (use as read only)
Public g_fov As Single                          'view frustrum field of view
Public g_aspect As Single                       'view frustrum aspect ratio
Public g_znear As Single                        'view frustrum near plane
Public g_zfar As Single                         'view frustrom far plane

                                                'Matrices (use as read only)
Public g_identityMatrix As D3DMATRIX            'Filled with Identity Matrix after D3DUtil_Init
Public g_worldMatrix As D3DMATRIX               'Filled with current world matrix
Public g_viewMatrix As D3DMATRIX                'Filled with current view matrix
Public g_projMatrix As D3DMATRIX                'Filled with current projection matrix

                                                'Clipplanes: use to ComputeClipPlanes to initialize
                                                'helpfull for view frustrum culling
Public g_ClipPlanes() As D3DPLANE               'Clipplane plane array
Public g_numClipPlanes As Long                  'Number of clip planes in g_ClipPlanes

Public light0 As D3DLIGHT8                      'light type usefull in imediate pane
Public light1 As D3DLIGHT8                      'light type usefull in imediate pane
  
Public g_TextureSampling As TextureParams       'defines how CreateTextureInPool sample textures

Public g_TextureLoadCallback  As Object         'object that implements LoadTextureCallback(sName as string) as Direct3dTexture8
Public g_bUseTextureLoadCallback As Boolean     'enables disables callback
  
Public g_mediaPath As String                    'Path to media and texture
                                                'read/write - must have ending backslash
                                                'best to use SetMediaPath to initialize



'------------------------------------------------------------------
'Global constants
'------------------------------------------------------------------

Public Const g_pi = 3.1415
Public Const g_InvertRotateKey = True   'flag to turn on fix for animation key problem
Public Const D3DFVF_VERTEX = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1

'------------------------------------------------------------------
'Locals
'------------------------------------------------------------------

'TexturePool Mangement data. see..
' D3DUTIL_LoadTextureIntoPool
' D3DUTIL_AddTextureToPool
' D3DUTIL_ReleaseTextureFromPool
' D3DUTIL_ReleaseAllTexturesFromPool
'
Dim m_texPool() As TexPoolEntry
Dim m_maxPool As Long
Dim m_nextEmpty As Long
Dim m_firstDel As Long

Const kGrowSize = 10


'------------------------------------------------------------------
'Functions
'------------------------------------------------------------------

'-----------------------------------------------------------------------------
'DOC: D3DUtil_SetupDefaultScene
'DOC:
'DOC: helper function that initializes some default lighting and render states
'DOC:
'DOC: remarks:
'DOC:   sets defaults for
'DOC:   g_fov, g_aspect, g_znear, g_zfar
'DOC:   g_identityMatrix, g_projMatrix, g_ViewMatrix, g_worldMatrix
'DOC:   set device state for project view and world matrices
'DOC:   set device state for 2 directional lights (0 and 1)
'DOC:   set device state for a default grey material
'-----------------------------------------------------------------------------

Public Sub D3DUtil_SetupDefaultScene()
    
    g_fov = g_pi / 4
    g_aspect = 1
    g_znear = 1
    g_zfar = 3000
    
    If g_lWindowHeight <> 0 And g_lWindowWidth <> 0 Then g_aspect = g_lWindowHeight / g_lWindowWidth
    
    D3DXMatrixIdentity g_identityMatrix
    
    D3DXMatrixPerspectiveFovLH g_projMatrix, g_fov, g_aspect, g_znear, g_zfar
    
    g_dev.SetTransform D3DTS_PROJECTION, g_projMatrix
    
    D3DXMatrixLookAtLH g_viewMatrix, vec3(0, 0, -20), vec3(0, 0, 0), vec3(0, 1, 0)
    
    g_dev.SetTransform D3DTS_VIEW, g_viewMatrix
                 
    g_dev.SetTransform D3DTS_WORLD, g_identityMatrix
    
    'default light0
    
    light0.Ambient = ColorValue4(1, 0.1, 0.1, 0.1)
    light0.diffuse = ColorValue4(1, 1, 1, 1)
    light0.Type = D3DLIGHT_DIRECTIONAL
    light0.Range = 10000
    light0.Direction.X = -1
    light0.Direction.Y = -1
    light0.Direction.z = -1
    D3DXVec3Normalize light0.Direction, light0.Direction
    g_dev.SetLight 0, light0
    g_dev.LightEnable 0, 1 'true
    
    'default light1
    
    light1.Ambient = ColorValue4(1, 0.3, 0.3, 0.3)
    light1.diffuse = ColorValue4(1, 1, 1, 1)
    light1.Type = D3DLIGHT_DIRECTIONAL
    light1.Range = 10000
    light1.Direction.X = 1
    light1.Direction.Y = -1
    light1.Direction.z = -1
    D3DXVec3Normalize light1.Direction, light1.Direction
    'g_dev.SetLight 1, light1
    'g_dev.LightEnable 1, 1 'true
        
        
    'set first material
    Dim material0 As D3DMATERIAL8
    material0.Ambient = ColorValue4(1, 0.2, 0.2, 0.2)
    material0.diffuse = ColorValue4(1, 0.5, 0.5, 0.5)
    material0.power = 10
    g_dev.SetMaterial material0
    
    With g_dev
        Call .SetRenderState(D3DRS_AMBIENT, &H10101010)
        Call .SetRenderState(D3DRS_CLIPPING, 1)             'CLIPPING IS ON
        Call .SetRenderState(D3DRS_LIGHTING, 1)             'LIGHTING IS ON
        Call .SetRenderState(D3DRS_ZENABLE, 1)              'USE ZBUFFER
        Call .SetRenderState(D3DRS_SHADEMODE, D3DSHADE_GOURAUD)
        
    End With
    
End Sub

'-----------------------------------------------------------------------------
'DOC: ColorValue4
'DOC: Params
'DOC:   a r g b   values valid between 0.0 and 1.0
'DOC: Return Value
'DOC:   a filled D3DCOLORVALUE type
'-----------------------------------------------------------------------------
Function ColorValue4(a As Single, r As Single, g As Single, b As Single) As D3DCOLORVALUE
    Dim c As D3DCOLORVALUE
    c.a = a
    c.r = r
    c.g = g
    c.b = b
    ColorValue4 = c
End Function

'-----------------------------------------------------------------------------
'DOC: Vec2
'DOC: Params
'DOC:   x y z   vector values
'DOC: Return Value
'DOC:   a filled D3DVECTOR type
'-----------------------------------------------------------------------------
Function vec2(X As Single, Y As Single) As D3DVECTOR2
    vec2.X = X
    vec2.Y = Y
End Function


'-----------------------------------------------------------------------------
'DOC: Vec3
'DOC: Params
'DOC:   x y z   vector values
'DOC: Return Value
'DOC:   a filled D3DVECTOR type
'-----------------------------------------------------------------------------
Function vec3(X As Single, Y As Single, z As Single) As D3DVECTOR
    vec3.X = X
    vec3.Y = Y
    vec3.z = z
End Function

'-----------------------------------------------------------------------------
'Name: FtoDW
'
'For calls that require that a single be packed into a long
'(such as some calls to SetRenderState) this function will do just that
'-----------------------------------------------------------------------------
Function FtoDW(f As Single) As Long
    Dim buf As D3DXBuffer
    Dim l As Long
    Set buf = g_d3dx.CreateBuffer(4)
    g_d3dx.BufferSetData buf, 0, 4, 1, f
    g_d3dx.BufferGetData buf, 0, 4, 1, l
    FtoDW = l
End Function

