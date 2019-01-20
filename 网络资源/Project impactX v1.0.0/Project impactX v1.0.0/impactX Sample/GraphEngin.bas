Attribute VB_Name = "xGraph"
'impactX Game Engine
'����ģ�����ڴ���DX�豸�ͼ��λ�ͼ
'ʹ�ñ���ģ���������:
'��������ʹ�ñ����漰����
'ʹ�ñ�������������ʹ���߳е�
'��������⿽����������룬�����뱣֤��������
'ϣ�����ܵõ���ʹ�ñ������������ĳ���
'Davy.xu sunicdavy@sina.com qq:20998333
Option Explicit
Dim dx As DirectX8
Dim D3D As Direct3D8
Public D3DDevice As Direct3DDevice8
Public D3DX As D3DX8
Dim D3DWindow As D3DPRESENT_PARAMETERS '��ʾģʽ�ĸ��ֲ���
Const TLFVF = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1 Or D3DFVF_SPECULAR

'Transforme and Lit�ṹ
Private Type TLVERTEX
  X As Single
  Y As Single
  Z As Single '2D��Ⱦʱ��ʹ��,��Ϊ0
  rhw As Single '��ʹ��,��Ϊ1
  Color As Long '������ɫ,AlphaBlend�Ļ�����ɫ
  Specular As Long '����,��ʹ��
  tu As Single '��ͼ����(0~1) x��
  tv As Single '��ͼ����(0~1) y��
End Type

'ͼ�δ���ʽ
Enum ENUM_XG_PROCESS
    xgSOFTWARE = D3DCREATE_SOFTWARE_VERTEXPROCESSING '���ģ��
    xgHARDWARE = D3DCREATE_HARDWARE_VERTEXPROCESSING 'Ӳ��ģ��
    xgPUREDEVICE = D3DCREATE_PUREDEVICE '��Ӳ��ģ��
    xgAUTO = 0 '�Զ�ʶ��(ȱʡ)
End Enum
'����͸������ɫ
Enum ENUM_XG_COLOR
    xgBLACK = &HFF000000 '��
    xgWHITE = &HFFFFFFFF '��
    
    xgRED = &HFFFF0000 '��
    xgGREEN = &HFF00FF00 '��
    xgBLUE = &HFF0000FF '��
    
    xgYELLOW = &HFFFFFF00 '��
    xgMAGENTA = &HFFFF00FF '���
    xgCYAN = &HFFFF00FF '��ɫ
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
'����
Dim xgMainFont As D3DXFont

    
'Fps ��ʾ
Private fpsTimer As Long
Private cCount As Integer
Private cFPS As Integer
'ʱ���ٶȿ���(LimitFPS)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private m_LastTime As Long
'����:��ʼ��DirectGraph
'����:ȫ��ģʽ�µ�ˮƽ�ʹ�ֱ�ֱ���,��ʾ���ڵľ��,D3D����ʽ
'ע��:��ʼ����,��Ļ����ɫ����ǿ�Ƶ���Ϊ16λɫ
Public Function InitDXGraph(ByVal ResWidth As Integer, ByVal ResHeight As Integer, ByVal hWnd As Long, Optional DisplayMode As ENUM_DISPLAYMODE = xgWindow, Optional Process As ENUM_XG_PROCESS) As Boolean
    On Error GoTo ErrH
    Dim BehaviorFlag As Long
    Dim caps As D3DCAPS8
    InitDXGraph = False
    Set dx = New DirectX8
    '�����趨D3D
    Set D3D = dx.Direct3DCreate()
    
    Set D3DX = New D3DX8
    
    Dim DMode As D3DDISPLAYMODE
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DMode
    If DisplayMode = xgWindow Then
        D3DWindow.Windowed = 1
        D3DWindow.SwapEffect = D3DSWAPEFFECT_DISCARD 'D3DSWAPEFFECT_COPY_VSYNC '
        D3DWindow.BackBufferFormat = DMode.Format
    Else
        'ҳ����(Frontbuffer��Backbuffer)���ܵ�ѡ��,����Ҫʹ��ȫ��Ļģʽ=D3DSWAPEFFECT_FLIP.
        D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
        '�����(Backbuffer)������
        D3DWindow.BackBufferCount = 1
        '����퓵Ć�λ�D�ظ�ʽ,D3DFMT_R5G6B5=16bits,65536ɫ,�߲�,
        '��Ȼ��Ҳ���԰����O����ȫ��(D3DFMT_R8G8B8).���҂����O������퓕r�䌍�����ڛQ���҂����@ʾģʽ
        D3DWindow.BackBufferFormat = D3DFMT_R5G6B5
            D3DWindow.EnableAutoDepthStencil = 1
        D3DWindow.AutoDepthStencilFormat = D3DFMT_D16
        'ˮƽ��ֱ�ֱ���
        D3DWindow.BackBufferWidth = ResWidth
        D3DWindow.BackBufferHeight = ResHeight
    End If
    
    '����window��hWnd
    D3DWindow.hDeviceWindow = hWnd
    If Process = xgAUTO Then
            D3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, caps
            '�Զ�ѡ����ʵ�Ӳ����ʾ�豸
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
    
    
    D3DDevice.SetVertexShader TLFVF '���VD3D���L��ʽʹ���҂��O���ķ�ʽ
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE '��ͼʱ����ALPHAͨ��(AlphaBlendʹ��)
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA '��Ⱦʱ����͸��ɫ
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA 'ͬ��
    
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
 SetTextType "SYSTEM", 11
InitDXGraph = True
 Exit Function
ErrH:
    Debug.Print "Err [InitDXGraph] ��ʼ��DX����"
    Debug.Print "�������ԭ��:�Կ���������ȷ,�Կ���֧��Direct3D,����Direct3D,�Դ治��"
    MsgBox "�������ԭ��:�Կ���������ȷ,�Կ���֧��Direct3D,����Direct3D,�Դ治��" & vbNewLine & "DirectX����:" & Hex(Err.Number), vbCritical, "��ʼ��DX����"
    End
End Function
'����:����Ļ�ϵ�ɫ
'����:����ͿĨ��Ļ����ɫ
Public Sub PaintScreen(ByVal BackColor As ENUM_XG_COLOR)
On Error GoTo ErrH
    If D3DDevice Is Nothing Then End
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, BackColor, 1#, 0
    Exit Sub
ErrH:
    Debug.Print "Err: [PaintScreen] �������ô���"
    End
End Sub
'����:�������,׼���豸
Public Sub RenderBegin()
On Error GoTo ErrH
    If D3DDevice Is Nothing Then End
    D3DDevice.BeginScene
    Exit Sub
ErrH:
    Debug.Print "Err: [RenderBegin] �������ô���,RenderBegin��RenderEnd�������ʹ��"
    End
End Sub
'����:������Ⱦ,����FPS,������ͼ�η�ת����ҳ��
Public Sub RenderEnd(Optional FrmInteval As Long = 1000)
On Error Resume Next
    If D3DDevice Is Nothing Then Exit Sub
    D3DDevice.EndScene '������Ⱦ
    '����FPS
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
    Debug.Print "Err: [RenderEnd] �������ô���,RenderBegin��RenderEnd�������ʹ��"
    End
End Sub
'����:����FPS
'����:��Ҫ��FPSֵ
Public Sub LimitFPS(Frame As Integer)
'֡ÿ��
    Do Until timeGetTime - m_LastTime > 1000 / Frame: Loop
    m_LastTime = timeGetTime
End Sub
'����:��ʼ������Ⱦ
'ע��:�뾡������������ַ���BeginText��EndText֮��
Public Sub BeginText()
    xgMainFont.Begin
End Sub
'����:����������Ⱦ
Public Sub EndText()
    xgMainFont.End
End Sub
'����:����Ļ�ϻ�������
'����:����,����,��ʾ��ɫ
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
    Debug.Print "Err [DrawText] ��������ʱ����"
    End
End Sub
'����:������ʾ����
'����:��������(��Word�����������ҵ�),�����С
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
    Debug.Print "Err [SetTextType] ��������ʱ����"
    End
End Sub
'D3D�ڲ�ʹ��,����һ��TLV����ṹ
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
'����:��һ����
'����:����������������,��ɫ
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
    Debug.Print "Err [DrawLine] ��ͼʱ����"
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
    Debug.Print "Err [DrawPoint] ��ͼʱ����"
    End
End Sub

'��Բ
Public Sub DrawCircle(X As Integer, Y As Integer, Radius As Integer, Color As ENUM_XG_COLOR)
    Dim iX As Integer, iY As Integer
    Static LastX As Integer, LastY As Integer
    Dim Angle As Single
    Dim tVer(0 To 1) As TLVERTEX
    Dim BlankTexture As Direct3DTexture8

    For Angle = 0 To 2 * 3.14 Step 3.14 / 30 '����Բ��Բ�ȣ�stepֵԽ��Խ��Բ
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

'����:��һ������
'����:���ε��ĸ���,��ɫ
Public Sub DrawRect(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, Color As ENUM_XG_COLOR)
    DrawLine Left, Top, Right, Top, Color
    DrawLine Right, Top, Right, Bottom, Color
    DrawLine Left, Bottom, Right, Bottom, Color
    DrawLine Left, Top, Left, Bottom, Color
End Sub
'����:��һ�����β����
'����:���ε��ĸ���,��ɫ
Public Sub DrawRectFill(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, Color As ENUM_XG_COLOR)
    On Error GoTo ErrH
    Dim tVer(0 To 3) As TLVERTEX
    Dim BlankTexture As Direct3DTexture8
    Dim t As Integer
    '�ߵ�����
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
    Debug.Print "Err [DrawRectFill] ��ͼʱ����"
    End
End Sub
'����:��һ��4ɫ������β����
'����:���ε��ĸ���,�ĸ�����ɫ
Public Sub DrawRectGradual(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, _
Color_LeftTop As ENUM_XG_COLOR, Color_LeftRight As ENUM_XG_COLOR, Color_BottomLeft As ENUM_XG_COLOR, Color_BottomRight As ENUM_XG_COLOR)
    On Error GoTo ErrH
    Dim tVer(0 To 3) As TLVERTEX
    Dim BlankTexture As Direct3DTexture8
    Dim t As Integer
    '�ߵ�����
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
    Debug.Print "Err [DrawRectFill] ��ͼʱ����"
    End
End Sub

'����:ж��DirectGraph
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
    D3DDevice.SetVertexShader TLFVF '���VD3D���L��ʽʹ���҂��O���ķ�ʽ
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE '��ͼʱ����ALPHAͨ��(AlphaBlendʹ��)
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA '��Ⱦʱ����͸��ɫ
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA 'ͬ��
    
End Sub
'���FPS
Public Property Get GetFPS() As Integer
    GetFPS = cFPS
End Property

