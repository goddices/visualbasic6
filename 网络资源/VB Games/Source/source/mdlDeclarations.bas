Attribute VB_Name = "mdlDeclarations"
Option Explicit

    '//API declarations
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long

    '//Pi...
    Public Const ConstPI As Single = 3.14159
    
    '//To convert to Radians
    Public Const RAD As Single = ConstPI / 180
    
    '//To convert to Degrees
    Public Const DEG As Single = 180 / ConstPI

    '//TL Vertex
    'Public Const D3DFVF_TLVERTEX = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)  'Or D3DFVF_SPECULAR Or D3DFVF_TEX1)

    '//Main DirectX8 object
'    Public ObjDX As DirectX8
'
'    '//Main Direct3D object
'    Public objD3D As Direct3D8
'
'    '//The Device used by Direct3D
'    Public g_dev As Direct3DDevice8
'
'    '//D3DX Helper library
'    Public ObjD3DX As D3DX8

    '//Public StartTime
    Public lngStartTime As Long
    
    Public myTexture As Direct3DTexture8
    
    '//TL Vertex
    Public Const D3DFVF_TLVERTEX = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
    
    Public Type typeTRANSLITVERTEX
        X As Single
        Y As Single
        z As Single
        rhw As Single
        color As Long
        tu As Single
        tv As Single
    End Type

Public Sub ResetStates()
    With g_dev
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
