VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGDISurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==========Class CGDISurface使用说明(RGB 5-5-5)============
'首先创建一个New CGDISurface.       (Dim x as New CGDISurface)
'CGDISurface.Read 文件名            (从文件名中载入图片,包括*.bmp,*.jpg,*.jif等)
'CGDISurface.Creat 图片宽,图片高    (在内存中创建一个空的 图片宽*图片高 24位位图)
'CGDISurface.ColorFilled R,G,B      (以指定的RGB值填充图象数据)
'CGDISurface.Clear  (清空颜色数据即ColorFilled黑色,比CGDISurface.ColorFilled 快)
'CGDISurface.DrawText X,Y,Text      (从(X,Y)处输出一个字符串)
'CGDISurface.Blt dx,dy,w,h,srcbmp,[sx],[sy]
'   dx,dy:即将在CGDISurface输出的左上坐标
'   w,h:即将在CGDISurface输出的图片宽和高
'   srcbmp:源图 CGDISurface类的实体,且必须已经有一个位图(Read 或 Creat 都可以)
'   sx,sy:源图的(sx,sy)处到(sx+w,sy+h)为处理区
'CGDISurface.LoadData Data() as byte
'   将图层的颜色数据复制到Data()数组,Data()为byte类型的动态数组
'   该过程调用后Data将设置成 (CGDISurface.pitch-1)*(CGDISurface.height-1)
'   的二维数组.注意!数据是上下颠倒的,所以图象上点(x,y)对应的数据为
'   Data(3*x,CGDISurface.height-1-y),宽为3*CGDISurface.width 应一次读取三个字节
'CGDISurface.Render dstdc           (将图形绘制到dstdc的设备环境)
'   eg:CGDISurface.Render Picture1.Hdc
'CGDISurface.hdc                    (DC设备句柄)
'CGDISurface.width                  (图象的宽)
'CGDISurface.height                 (图象的高)
'CGDISurface.pitch             (一行图象数据的字节数)
'CGDISurface.lpData                 (指向图象数据区首地址的指针)

Private Type BITMAPFILEHEADER
    betype As Integer               ' 图 象 文 件 型 态
    bfsize As Long                  ' 图 象 文 件 大 小
    bfreserved1 As Integer          ' 保 留 位
    bfreserved2 As Integer          ' 保 留 位
    bfoffbits As Long               ' 图 象 数 据 的 偏 移 量
End Type


Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type


Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Type RECT
    Left As Integer
    Right As Integer
    Top As Integer
    Bottom As Integer
End Type

Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal HDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetObjectA Lib "gdi32" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Private m_ptr As Long                                   '图象的数据区指针
Private m_hdc As Long                                   '图象的设备环境
Private m_width As Integer
Private m_height As Integer
Private m_pitch As Integer

Private m_BmpHeader As BITMAPINFOHEADER
Private m_BmpInfo As BITMAPINFO

Private Sub Class_Terminate()
    DeleteDC m_hdc
End Sub

Public Property Get HDC() As Long
    HDC = m_hdc
End Property

Public Property Get lpData() As Long
    lpData = m_ptr
End Property

Public Property Get Width() As Integer
    Width = m_width
End Property

Public Property Get Height() As Integer
    Height = m_height
End Property

Public Property Get pitch() As Integer
    pitch = m_pitch
End Property

Public Function Create(ByVal Width As Integer, ByVal Height As Integer) As Boolean
    Dim hdib As Long
    If m_hdc <> 0 Then DeleteDC m_hdc
    
    m_hdc = CreateCompatibleDC(0)
    
    If m_hdc = 0 Then Exit Function
    
    With m_BmpInfo.bmiHeader
        .biSize = Len(m_BmpInfo.bmiHeader)
        .biWidth = Width
        .biHeight = -Height '注意使用了负值
        .biPlanes = 1
        .biBitCount = 16
        .biCompression = 0 'BI_RGB
        m_pitch = ((.biWidth * 2 + 2) And &HFFFFFFFC)
        .biSizeImage = m_pitch * .biHeight
    End With
    
    m_BmpHeader = m_BmpInfo.bmiHeader         '===
    
    hdib = CreateDIBSection(m_hdc, m_BmpInfo, 0, m_ptr, 0, 0)
    SelectObject m_hdc, hdib
    DeleteObject hdib
    
    m_width = Width
    m_height = Height
    Clear
    
    Create = True
End Function

Public Function Read(ByVal FileName As String) As Boolean
    Dim tpic As Picture
    Dim bmp As BITMAP
    Dim HDC As Long

    On Error GoTo errhand:
    
    Set tpic = LoadPicture(FileName)
    If tpic Is Nothing Then GoTo errhand
    
    Call GetObjectA(tpic.handle, Len(bmp), bmp)
    '创建一个bmp.bmWidth* bmp.bmHeight大小空的24位图
    If Not Create(bmp.bmWidth, bmp.bmHeight) Then Exit Function
    
    HDC = CreateCompatibleDC(0)
    Call SelectObject(HDC, tpic.handle)
    BitBlt m_hdc, 0, 0, bmp.bmWidth, bmp.bmHeight, HDC, 0, 0, vbSrcCopy
    DeleteDC HDC
    
    Set tpic = Nothing
    Read = True
    Exit Function
    
errhand:
    Read = False
    Err.Raise vbObjectError + 513, "VBGameEngine", "Picture File Load Failed!"
End Function

Public Sub LoadData(data() As Byte, Optional Style = 2)
    If Style = 2 Then
        ReDim data(m_pitch - 1, m_height - 1)
        CopyMemory data(0, 0), ByVal m_ptr, CLng(m_pitch) * m_height
    ElseIf Style = 1 Then
        ReDim data(CLng(m_pitch) * m_height - 1)
        CopyMemory data(0), ByVal m_ptr, CLng(m_pitch) * m_height
    End If
End Sub

Public Sub SetData(data() As Byte, Optional Style = 2)
    If Style = 2 Then
        CopyMemory ByVal m_ptr, data(0, 0), CLng(m_pitch) * m_height
    ElseIf Style = 1 Then
        CopyMemory ByVal m_ptr, data(0), CLng(m_pitch) * m_height
    End If
End Sub

Public Sub Render(dstdc As Long)
    If m_ptr = 0 Then Exit Sub
    BitBlt dstdc, 0, 0, m_width, m_height, m_hdc, 0, 0, vbSrcCopy
End Sub

Public Sub Blt(ByVal dx As Integer, ByVal dy As Integer, ByVal W As Integer, ByVal H As Integer _
    , srcBmp As CGDISurface, Optional ByVal sx As Integer, Optional ByVal sy As Integer)
    'On Error GoTo errhand
    If m_ptr = 0 Then Exit Sub
    
    If Not Clipper(dx, dy, W, H, sx, sy, srcBmp.Width, srcBmp.Height) Then Exit Sub

    BitBlt m_hdc, dx, dy, W, H, srcBmp.HDC, sx, sy, vbSrcCopy
    Exit Sub
    
errhand:
    Err.Clear
End Sub

Public Sub Clear()
    ZeroMemory ByVal m_ptr, CLng(m_pitch) * m_height
End Sub

Public Sub ColorFilled(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)
    '暂时不支持！
    'MemorySet ByVal m_ptr, DDRGB(R, G, B), CLng(m_pitch) * m_height
End Sub

Public Function GetPixel(ByVal x As Integer, ByVal y As Integer) As Integer
    'Debug.Assert x < m_pitch / 2 And y < m_height And x >= 0 And y >= 0
    
    If x < m_pitch / 2 And y < m_height And x >= 0 And y >= 0 Then
        CopyMemory GetPixel, ByVal m_ptr + y * CLng(m_pitch) + x * 2, 2
    Else
        GetPixel = -1
    End If
End Function

Public Sub DrawText(ByVal x As Integer, ByVal y As Integer, ByVal s As String)
    If m_hdc = 0 Then Exit Sub
    TextOut m_hdc, x, y, s, LenB(StrConv(s, vbFromUnicode))
End Sub

Private Function Clipper(x As Integer, y As Integer, W As Integer, H As Integer, sx As Integer, sy As Integer, sw As Integer, sh As Integer) As Boolean
    '自动管理区域裁剪
    If x + W > m_width - 1 Then
        W = m_width - x - 1
        If W < 0 Then Exit Function
    End If
    If y + H > m_height - 1 Then
        H = m_height - y - 1
        If H < 0 Then Exit Function
    End If
    
    If x < 0 Then
        W = W + x
        sx = sx - x
        x = 0
    End If
    If y < 0 Then
        H = H + y
        sy = sy - y
        y = 0
    End If
    
    If sx + W > sw - 1 Then
        W = sw - sx - 1
        If W < 0 Then Exit Function
    End If
    If sy + H > sh - 1 Then
        H = sh - sy - 1
        If H < 0 Then Exit Function
    End If

    If sx < 0 Then
        W = W + sx
        sx = 0
    End If
    If sy < 0 Then
        H = H + sy
        sy = 0
    End If
    
    Clipper = True
End Function
