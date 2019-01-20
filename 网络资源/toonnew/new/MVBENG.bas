Attribute VB_Name = "MVBENG"

Global dx As DirectX8
Global D3DX As D3DX8

Global D3D As Direct3D8
Global device As Direct3DDevice8

Global mat As D3DMATRIX
Global matWorld As D3DMATRIX
Global matView As D3DMATRIX
Global matProj As D3DMATRIX
Global Const rad1 = 3.14 / 180
Global D3DWindow As D3DPRESENT_PARAMETERS

Global d3dsdBackBuffer As D3DSURFACE_DESC

Type OneMesh
  numX As Long
  meshX As D3DXMesh
  mateX() As D3DMATERIAL8
  nIndici As Long
  nFacce As Long
End Type

Global CosRx As Single
Global CosRy As Single
Global CosRz As Single
Global SinRx As Single
Global SinRy As Single
Global SinRz As Single

Sub creaScen(dxWidth As Long, dxHeight As Long, dxBpp As CONST_D3DFORMAT, Fhwnd As Long, finestra As Boolean, numBackBuffer As Long)
  Dim DispMode As D3DDISPLAYMODE
  '
  Set D3D = dx.Direct3DCreate()
  D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

  If finestra Then
    D3DWindow.Windowed = 1
    D3DWindow.BackBufferCount = 1
    D3DWindow.BackBufferFormat = DispMode.Format
  Else
    D3DWindow.BackBufferCount = numBackBuffer
    D3DWindow.BackBufferFormat = dxBpp
    D3DWindow.BackBufferWidth = dxWidth
    D3DWindow.BackBufferHeight = dxHeight
  End If

  D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
  D3DWindow.EnableAutoDepthStencil = 1
  D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
  D3DWindow.hDeviceWindow = Fhwnd
  D3DWindow.flags = D3DPRESENTFLAG_LOCKABLE_BACKBUFFER

  Set device = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Fhwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)

  Dim pBackBuffer As Direct3DSurface8
  Set pBackBuffer = device.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
  pBackBuffer.GetDesc d3dsdBackBuffer

End Sub

Public Function muovMatrix(scala As Single, angleX As Long, angleY As Long, angleZ As Long, posX As Single, posY As Single, posZ As Single) As D3DMATRIX
  CosRx = Cos(angleX * rad1)
  CosRy = Cos(angleY * rad1)
  CosRz = Cos(angleZ * rad1)
  SinRx = Sin(angleX * rad1)
  SinRy = Sin(angleY * rad1)
  SinRz = Sin(angleZ * rad1)

  With muovMatrix
    .m11 = (scala * CosRy * CosRz)
    .m12 = (scala * CosRy * SinRz)
    .m13 = -(scala * SinRy)

    .m21 = -(scala * CosRx * SinRz) + (scala * SinRx * SinRy * CosRz)
    .m22 = (scala * CosRx * CosRz) + (scala * SinRx * SinRy * SinRz)
    .m23 = (scala * SinRx * CosRy)

    .m31 = (scala * SinRx * SinRz) + (scala * CosRx * SinRy * CosRz)
    .m32 = -(scala * SinRx * CosRz) + (scala * CosRx * SinRy * SinRz)
    .m33 = (scala * CosRx * CosRy)

    .m41 = posX
    .m42 = posY
    .m43 = posZ
    .m44 = 1#
  End With
End Function

Function creaTex(filesrc As String, ColorKey As Long, Optional coloreK As Boolean = False) As Direct3DTexture8
  If coloreK Then
    Set creaTex = D3DX.CreateTextureFromFileEx(device, filesrc, D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorKey, ByVal 0, ByVal 0)
  Else
    Set creaTex = D3DX.CreateTextureFromFile(device, filesrc)
  End If
End Function

Function creaMesh(filesrc As String, Optional texture As Boolean = False, Optional materiali As Boolean = False, Optional directory As String) As OneMesh

  Dim MtrlBuffer As D3DXBuffer
  Set creaMesh.meshX = D3DX.LoadMeshFromX(filesrc, D3DXMESH_SYSTEMMEM, device, Nothing, MtrlBuffer, creaMesh.numX)
  Dim retAdjacency As D3DXBuffer
  creaMesh.nIndici = creaMesh.meshX.GetNumVertices
  creaMesh.nFacce = creaMesh.meshX.GetNumFaces
  ReDim creaMesh.mateX(creaMesh.numX)
  Dim strTexName As String
  For i = 0 To creaMesh.numX - 1
    If materiali Then
      D3DX.BufferGetMaterial MtrlBuffer, i, creaMesh.mateX(i)
    End If
  Next
  Set MtrlBuffer = Nothing

End Function

Function creaFont(f As StdFont) As D3DXFont
  Dim iT As IFont
  Set iT = f
  Set creaFont = D3DX.CreateFont(device, iT.hFont)
End Function

Sub termina(Optional spegni As Boolean = True)
  On Error Resume Next

   Set cubo.meshX = Nothing
  Erase cubo.mateX
  Set T = Nothing

  Dim emptypresent As D3DPRESENT_PARAMETERS
  D3DWindow = emptypresent
  If spegni Then End
  End Sub
