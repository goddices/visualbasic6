Attribute VB_Name = "drawpic"
'绘制水墨线
Option Explicit

Public Const FVF_TLV = (D3DFVF_XYZRHW Or D3DFVF_TEX1)
Public Type CUSTOMVERTEXVerte
 P As D3DVECTOR
 rhw As Single
 T As D3DVECTOR2
End Type

Dim kNumsmo As Long
Dim smo() As D3DVECTOR
Dim Vertices() As CUSTOMVERTEXVerte
Dim Linklen() As Single
Dim LinklenMax As Single

Dim i As Long, j As Long, K As Long

Dim Vt1 As D3DVECTOR, Vt2 As D3DVECTOR
Dim Vtn1 As D3DVECTOR, Vtn2 As D3DVECTOR

Public Sub SwapChain(LiDate() As D3DVECTOR)
 Dim i As Long, LS As Long
 LS = UBound(LiDate)
 Dim LiTMP() As D3DVECTOR
 ReDim LiTMP(LS)
 For i = 0 To LS
  LiTMP(i) = LiDate(LS - i)
 Next
 LiDate = LiTMP
End Sub

Public Sub SplitChain(LiDate() As D3DVECTOR) '细分（平滑用）
 Dim i As Long, LS As Long
 LS = UBound(LiDate)
 Dim LiTMP() As D3DVECTOR
 ReDim LiTMP(LS * 2 - 1)
 LiTMP(0) = LiDate(0): LiTMP(LS * 2 - 1) = LiDate(LS)
 For i = 1 To LS - 1
  LiTMP(2 * i - 1) = VParam(LiDate(i), LiDate(i - 1), 0.3)
  LiTMP(2 * i) = VParam(LiDate(i), LiDate(i + 1), 0.3)
 Next
 LiDate = LiTMP
End Sub

Public Sub drawChain(Li() As D3DVECTOR)
 LinklenMax = 0
 kNumsmo = UBound(Li)
 If kNumsmo < 2 Then Exit Sub
 ReDim smo(kNumsmo)

 For i = 0 To kNumsmo
  smo(i) = Li(i)
 Next

 If Abs(smo(0).x - smo(kNumsmo).x) > Abs(smo(0).y - smo(kNumsmo).y) Then
  If smo(0).x > smo(kNumsmo).x Then SwapChain smo()
 Else
  If smo(0).y > smo(kNumsmo).y Then SwapChain smo()
 End If

 'SplitChain smo()
 kNumsmo = UBound(smo)

 ReDim Linklen(kNumsmo - 1)
 For i = 0 To kNumsmo - 1
  Linklen(i) = VDst(smo(i + 1), smo(i))
  LinklenMax = LinklenMax + Linklen(i)
 Next

 Dim LinklenINC As Single
 Dim VNor As D3DVECTOR

 ReDim Vertices(2 * kNumsmo + 1)
 For i = 0 To kNumsmo
  If i = 0 Then
   VNor = caleNor(smo(1), smo(0))
  ElseIf i = kNumsmo Then
   VNor = caleNor(smo(kNumsmo), smo(kNumsmo - 1))
  Else
   Vtn1 = caleNor(smo(i + 1), smo(i))
   Vtn2 = caleNor(smo(i), smo(i - 1))
   VNor = Normalize(Add(Vtn1, Vtn2))
   VNor = VScale(VNor, 1 / Dot(VNor, Vtn1))
  End If
  VNor = VScale(VNor, 5)
  Vertices(i * 2).P = Subtract(smo(i), VNor)
  Vertices(i * 2 + 1).P = Add(smo(i), VNor)
  Vertices(i * 2).rhw = 1
  Vertices(i * 2 + 1).rhw = 1
  Vertices(i * 2).T = vec2(LinklenINC, 0)
  Vertices(i * 2 + 1).T = vec2(LinklenINC, 1)
  If i <> kNumsmo Then LinklenINC = LinklenINC + Linklen(i) / LinklenMax
 Next

 device.SetRenderState D3DRS_ALPHABLENDENABLE, 1
 device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
 device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

 device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG1
 device.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
 device.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
 device.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE

 device.SetTexture 0, T
 device.SetVertexShader FVF_TLV
 device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2 * kNumsmo, Vertices(0), Len(Vertices(0))
End Sub

Public Function caleNor(starver As D3DVECTOR, endver As D3DVECTOR) As D3DVECTOR
 caleNor = Normalize(cross(Subtract(starver, endver), vec3(0, 0, 1)))
End Function

Public Function GetXYScreen(PNT As D3DVECTOR, Optional Xpos As Long, Optional YPOS As Long) As D3DVECTOR
 Dim Vtmp As D3DVECTOR4, vout As D3DVECTOR4
 Vtmp = vec3Tovec4(PNT)
 D3DXVec4Transform vout, Vtmp, mat
 If vout.w = 0 Then vout.w = 0.00001
 Xpos = (vout.x / vout.w + 1) / 2 * (d3dsdBackBuffer.Width)
 YPOS = (1 - vout.y / vout.w) / 2 * (d3dsdBackBuffer.Height)
 GetXYScreen.x = Xpos: GetXYScreen.y = YPOS: GetXYScreen.Z = 0
End Function
