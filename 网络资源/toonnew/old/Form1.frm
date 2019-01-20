VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "China Render  ZH1110"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   505
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   783
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   7335
      Left            =   240
      ScaleHeight     =   489
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   745
      TabIndex        =   0
      Top             =   120
      Width           =   11175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'连续边界的水墨渲染

'天恩软件

'——2005.4.11
'by: zh1110

Const FVF_VERTEX2 = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)

Dim shadow As ombra
Dim testo As D3DXFont

Dim pntx As Single, pnty As Single

Dim body As Long
Dim shver2 As Long
Dim angoloX As Long
Dim angoloY As Long
Dim memoX As Long
Dim memoY As Long
Dim memos As Long
Public scal As Single
Dim premuto As Boolean '

Private Sub Form_Load()
  D = 0.5
  scal = 10
  creaScen 800, 600, D3DFMT_R5G6B5, Picture1.hWnd, 1, 1

  Form1.Show
  Set testo = creaFont(Form1.Font)
  STAR

  Set T = creaTex(App.Path & "\LinkPic2.TGA", 0, False)
  'Set T = creaTex(App.Path & "\LinkPic.TGA", 0, False)

  device.SetRenderState D3DRS_LIGHTING, 0
  device.SetRenderState D3DRS_SHADEMODE, D3DSHADE_GOURAUD
  device.SetRenderState D3DRS_ZENABLE, 1

  D3DXMatrixPerspectiveFovLH matProj, 45 * rad1, 4 / 5, 0.1, 1400
  device.SetTransform D3DTS_PROJECTION, matProj

  Call mainLoop
End Sub

Sub mainLoop()
1
  drawse
  DoEvents
  GoTo 1
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    If X > memoX Then angoloX = angoloX - Abs(X - memoX) / 1
    If X < memoX Then angoloX = angoloX + Abs(X - memoX) / 1
    If Y > memoY Then angoloY = angoloY + Abs(Y - memoY) / 1
    If Y < memoY Then angoloY = angoloY - Abs(Y - memoY) / 1
  ElseIf Button = 2 Then
    scal = scal + (Y - memos) / 30
    If scal > 44 Then scal = 44
    If scal < 2 Then scal = 2
  Else
    premuto = False
  End If
  memoX = X
  memoY = Y
  memos = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Unload Me
  termina True
End Sub

Private Sub drawse()
  Dim i As Long

  device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, D3DColorMake(0.8, 0.8, 0.8, 0), 1, 0
  device.BeginScene

  device.SetRenderState D3DRS_ALPHABLENDENABLE, 0
  device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG1
  device.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_DIFFUSE
  device.SetRenderState D3DRS_CULLMODE, 1

  matWorld = muovMatrix(scal / objra, -angoloY, angoloX, 0, 12, 0, 0)
  device.SetTransform D3DTS_WORLD, matWorld
  D3DXMatrixLookAtLH matView, vec3(0, 0, -40), vec3(0, 0, 0), vec3(0, 1, 0)
  device.SetTransform D3DTS_VIEW, matView

  For i = 0 To cubo.numX - 1
    cubo.meshX.DrawSubset aus
  Next i
  Dim m As D3DMATRIX
  Dim DVerts(1) As CUSTOMVERCOL
  Dim Vtn1 As D3DVECTOR, Vtn2 As D3DVECTOR
  D3DXMatrixMultiply m, matWorld, matView

  ReDim Varray(0)

  For i = 0 To (UBound(shadow.vertici) - 1) / 2
    DVerts(0).P = shadow.vertici(i * 2).postion
    DVerts(1).P = shadow.vertici(i * 2 + 1).postion
    D3DXVec3TransformNormal Vtn1, shadow.vertici(i * 2).normal1, m
    D3DXVec3TransformNormal Vtn2, shadow.vertici(i * 2).normal2, m
    If Vtn1.Z * Vtn2.Z < 0 Then Call AddToArray(DVerts)
  Next

  D3DXMatrixLookAtLH matView, vec3(0, 0, -39.8), vec3(0, 0, 0), vec3(0, 1, 0)
  device.SetTransform D3DTS_VIEW, matView '测试可见点（黑色）
  D3DXMatrixMultiply mat, matWorld, matView
  D3DXMatrixMultiply mat, mat, matProj
  Dim V1 As D3DVECTOR4, vout As D3DVECTOR4

  device.SetVertexShader D3DFVF_CUSTOMVERCOL
  For i = 1 To UBound(Varray)
    DVerts(0).P = Vmid(Varray(i).P1, Varray(i).P2)
    device.DrawPrimitiveUP D3DPT_POINTLIST, 1, DVerts(0), Len(DVerts(0))
  Next

  For i = 1 To UBound(Varray)
    If i > UBound(Varray) Then Exit For
    DVerts(0).P = Varray(i).P1
    DVerts(1).P = Varray(i).P2
    GetScreen VScale(Add(DVerts(0).P, DVerts(1).P), 0.5), pntx, pnty

    'If GetPixel(Picture1.hdc, pntx, pnty) <> 0 Then
    If Not (GetPixel(Picture1.hdc, pntx, pnty) = 0 Or _
       GetPixel(Picture1.hdc, pntx - 1, pnty) = 0 Or _
       GetPixel(Picture1.hdc, pntx, pnty - 1) = 0) Then
      'Picture1.Circle (pntx, pnty), 2, QBColor(8): Picture1.Print ii
      Call DelArray(DVerts)
      i = i - 1
    End If
  Next

  Call FindAndDrawChian
  matlast = mat

  Dim r As RECT
  r = MakeRect(800, 20, 20, 600)
  testo.DrawTextW "BY zh1110 " & "  Out Chain：" & (UBound(shadow.vertici) + 1) / 2, -1, r, DT_LEFT Or DT_WORDBREAK, D3DColorMake(1, 1, 1, 1)

  device.EndScene
  device.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Public Sub STAR()
  cubo = creaMesh(App.Path & mane, True, True, App.Path)
  Set cubo.meshX = cubo.meshX.CloneMeshFVF(D3DXMESH_MANAGED, FVF_VERTEX2, device)
  D3DX.ComputeBoundingSphereFromMesh cubo.meshX, objcen, objra
  shadow = CreaOutLine(cubo.meshX)
End Sub
