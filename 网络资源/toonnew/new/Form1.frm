VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "China Render  ZH1110"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11730
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   782
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   9120
      Left            =   120
      ScaleHeight     =   608
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   768
      TabIndex        =   0
      Top             =   120
      Width           =   11520
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'连续边界的水墨渲染(BSP)

'天恩软件

'——2005.4.11
'by: zh1110

Dim shadow As silLine
Dim m_font As D3DXFont

Dim body As Long
Dim shver2 As Long
Dim angoloX As Long
Dim angoloY As Long
Dim memoX As Long
Dim memoY As Long
Dim memos As Long
Public scal As Single
Dim premuto As Boolean '

Sub mainLoop()
  ReDim Varray((UBound(shadow.vertici) + 1) / 2)
  ReDim Varray2((UBound(shadow.vertici) + 1) / 2) 'Varray2是中间变量

  D3DXMatrixLookAtLH matView, vec3(0, 0, -40), vec3(0, 0, 0), vec3(0, 1, 0)
  device.SetTransform D3DTS_VIEW, matView
1
  drawse
  DoEvents
  GoTo 1
End Sub

Private Sub Form_Load()
  D = 0.5
  scal = 16

  If dx Is Nothing Then Set dx = New DirectX8
  If D3DX Is Nothing Then Set D3DX = New D3DX8

  creaScen 800, 600, D3DFMT_R5G6B5, Picture1.hWnd, 1, 1
  Set m_font = creaFont(Form1.Font)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 1 Then
    If x > memoX Then angoloX = angoloX - Abs(x - memoX) / 1
    If x < memoX Then angoloX = angoloX + Abs(x - memoX) / 1
    If y > memoY Then angoloY = angoloY + Abs(y - memoY) / 1
    If y < memoY Then angoloY = angoloY - Abs(y - memoY) / 1
  ElseIf Button = 2 Then
    scal = scal + (y - memos) / 30
    If scal > 44 Then scal = 44
    If scal < 2 Then scal = 2
  Else
    premuto = False
  End If
  memoX = x
  memoY = y
  memos = y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set D3D = Nothing
  Set device = Nothing
  Set m_font = Nothing
  Set D3DX = Nothing
  Set dx = Nothing
  Unload Me
  termina True
End Sub

Private Sub drawse()
  Dim i As Long

  device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, D3DColorMake(0.9, 0.9, 0.9, 0), 1, 0
  device.BeginScene

  Dim mtrl As D3DMATERIAL8
  mtrl.diffuse = MCOLOR(1, 1, 1, 0)
  mtrl.Ambient = mtrl.diffuse
  device.SetMaterial mtrl

  device.SetRenderState D3DRS_LIGHTING, 1
  device.SetRenderState D3DRS_ALPHABLENDENABLE, 0
  device.SetRenderState D3DRS_CULLMODE, 1
  device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG1
  device.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_DIFFUSE
  device.SetRenderState D3DRS_AMBIENT, D3DColorMake(0.9, 0.9, 0.9, 0)

  matWorld = muovMatrix(scal / objra, -angoloY, angoloX, 0, 0, -2, 0)
  device.SetTransform D3DTS_WORLD, matWorld
  D3DXMatrixPerspectiveFovLH matProj, 45 * rad1, 4 / 5, 10, 1400
  device.SetTransform D3DTS_PROJECTION, matProj

  For i = 0 To cubo.numX - 1
    cubo.meshX.DrawSubset i
  Next i
  device.SetRenderState D3DRS_LIGHTING, 0

  arraylen = 0
  Dim m As D3DMATRIX
  D3DXMatrixMultiply m, matWorld, matView

  Dim Pc1 As D3DVECTOR, Pc2 As D3DVECTOR
  Dim Vtn1 As D3DVECTOR, Vtn2 As D3DVECTOR, Vpt As D3DVECTOR4
  For i = 0 To (UBound(shadow.vertici) - 1) / 2
    Pc1 = shadow.vertici(i * 2).postion
    Pc2 = shadow.vertici(i * 2 + 1).postion
    D3DXVec3TransformNormal Vtn1, shadow.vertici(i * 2).normal1, m
    D3DXVec3TransformNormal Vtn2, shadow.vertici(i * 2).normal2, m
    D3DXVec3Transform Vpt, Pc1, m
    If Dot43(Vpt, Vtn1) * Dot43(Vpt, Vtn2) < 0 Then Call AddToArray(Pc1, Pc2)
  Next

  D3DXMatrixPerspectiveFovLH matProj, 45 * rad1, 4 / 5, 10.1, 1600 '适当减小z缓冲值
  device.SetTransform D3DTS_PROJECTION, matProj '测试可见点（黑色）
  D3DXMatrixMultiply mat, matWorld, matView
  D3DXMatrixMultiply mat, mat, matProj

  Dim DVerts() As CUSTOMVERCOL
  ReDim DVerts(arraylen)
  device.SetVertexShader D3DFVF_CUSTOMVERCOL
  For i = 1 To arraylen
    DVerts(i).P = Vmid(Varray(i).P1, Varray(i).P2)
  Next
  device.DrawPrimitiveUP D3DPT_POINTLIST, arraylen, DVerts(0), Len(DVerts(0))

  For i = 1 To arraylen
    Varray2(i) = Varray(i)
    Varray2(i).use = 0 '清除标志
  Next

  '锁住背景
  Dim pData As D3DLOCKED_RECT
  Dim lp As Long, pD As Long
  Dim pntx As Long, pnty As Long
  Dim pBackBuffer As Direct3DSurface8
  Set pBackBuffer = device.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
  pBackBuffer.LockRect pData, ByVal 0, 0

If d3dsdBackBuffer.Format = D3DFMT_X8R8G8B8 Or d3dsdBackBuffer.Format = D3DFMT_A8R8G8B8 Then '32色
  For i = 1 To arraylen
    Call GetXYScreen(DVerts(i).P, pntx, pnty)
    lp = pData.pBits + pData.Pitch * pnty + 4 * pntx
    DXCopyMemory pD, ByVal lp, 3 '检查颜色位，X或A不检查
    If pD <> 0 Then Varray2(i).use = True
  Next
Else                                            '16色
  For i = 1 To arraylen
    Call GetXYScreen(DVerts(i).P, pntx, pnty)
    lp = pData.pBits + pData.Pitch * pnty + 2 * pntx
    DXCopyMemory pD, ByVal lp, 2
    If pD <> 0 Then Varray2(i).use = True
  Next
End If

  pBackBuffer.UnlockRect

  Dim leng As Long
  leng = arraylen
  arraylen = 0
  For i = 1 To leng
    If Varray2(i).use = 0 Then
      arraylen = arraylen + 1
      Varray(arraylen) = Varray2(i)
    End If
  Next

  For i = 1 To arraylen '转化到屏幕坐标(平面)
    Varray(i).P1 = GetXYScreen(Varray(i).P1)
    Varray(i).P2 = GetXYScreen(Varray(i).P2)
  Next

  device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, D3DColorMake(0.9, 0.9, 0.9, 0), 1, 0
  Call FindAndDrawChian

  Dim r As RECT
  r = MakeRect(40, 10, 20, 200)
  m_font.DrawTextW "Vertices：" & (UBound(shadow.vertici) + 1) / 2, -1, r, DT_LEFT Or DT_WORDBREAK, D3DColorMake(1, 0, 0, 1)

  device.EndScene
  device.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Public Sub STAR()
  termina 0

  Set T = creaTex(App.Path & penmane, 0, False)

  Const FVF_VERTEX2 = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)
  cubo = creaMesh(App.Path & mane, True, True, App.Path)
  Set cubo.meshX = cubo.meshX.CloneMeshFVF(D3DXMESH_MANAGED, FVF_VERTEX2, device)
  D3DX.ComputeBoundingSphereFromMesh cubo.meshX, objcen, objra
  shadow = CreaOutLine(cubo.meshX)

  device.SetRenderState D3DRS_LIGHTING, 0
  device.SetRenderState D3DRS_SHADEMODE, D3DSHADE_GOURAUD
  device.SetRenderState D3DRS_ZENABLE, 1
  D3DXMatrixPerspectiveFovLH matProj, 45 * rad1, 4 / 5, 0.1, 1400
  device.SetTransform D3DTS_PROJECTION, matProj

  Call mainLoop

End Sub
