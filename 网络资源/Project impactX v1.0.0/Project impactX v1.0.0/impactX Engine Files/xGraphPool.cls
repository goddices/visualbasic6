VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xGraphPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'impactX Game Engine
'本类模块用于存储图像的各种数据和图像的操作,必须与xGraph配合使用
'使用本类模块必须遵守:
'你可以免费使用本引擎及代码
'使用本引擎后的责任由使用者承担
'你可以任意拷贝本引擎代码，但必须保证其完整性
'希望我能得到你使用本引擎制作出的程序
'Davy.xu sunicdavy@sina.com qq:20998333
Option Explicit
Private Type impactGraph
    Texture As Direct3DTexture8
    Width As Integer
    Height As Integer
    SourceRect As RECT '图像显示范围
    Translation As D3DVECTOR2 '图像显示坐标
    Scale As D3DVECTOR2 '缩放
    Center As D3DVECTOR2 '图像中心
    Rotation As Single '旋转角度
    Color As Long '颜色(半透明,滤色镜)
    Hnum As Integer '水平分割数(精灵动画)
    Vnum As Integer '垂直分割数
    CellIndex As Single  '显示精灵分块索引号
    MirrorType As ENUM_XG_MIRROR
    D3DDevice As Direct3DDevice8
    D3Dobj As D3DX8
    FrmSpd As Long
End Type

Private Type ImageSize
     Width As Long
     Height As Long      '圖形的高度
End Type

Enum ENUM_XG_MIRROR
    xgNONE = 0
    xgH_MIRROR = 1
    xgV_MIRROR = 2
End Enum
Dim Sprite As D3DXSprite
Dim gfx As impactGraph '图形的各种信息
Dim MirrorMatrix As D3DMATRIX
Dim pxArr() As Byte
Dim isMirrorOpen As Boolean
'功能: 载入图片
'参数:图片索引号,路径(不带盘符时默认为工作目录),透明色
Public Function LoadGraph(ByVal Pathname As String, ByVal ColorKey As ENUM_XG_COLOR, Optional Hnum As Integer = 1, Optional Vnum As Integer = 1) As Boolean
    'On Error GoTo ErrH
    If Len(Pathname) = 0 Then
        Debug.Print "Err: [LoadGraph] 路径不能为空"
        End
    Else
        '将默认文件名变为工作目录
        If Mid(Pathname, 2, 1) <> ":" Then Pathname = App.Path & "\" & Pathname
        If Len(Dir(Pathname)) = 0 Then
            Debug.Print "Err: [LoadGraph] 文件不存在"
            Debug.Print "     Pathname:" & Pathname
            End
        End If
    End If
    Pathname = Trim(Pathname) '去空格
    Dim GfxInfo As D3DXIMAGE_INFO
    With gfx
        Set .Texture = gfx.D3Dobj.CreateTextureFromFileEx( _
                                                    .D3DDevice, _
                                                    Pathname, _
                                                    D3DX_DEFAULT, _
                                                    D3DX_DEFAULT, _
                                                    1, _
                                                    0, D3DFMT_UNKNOWN, _
                                                    D3DPOOL_MANAGED, _
                                                    D3DX_FILTER_NONE, _
                                                    D3DX_FILTER_NONE, _
                                                    ColorKey, ByVal 0, ByVal 0)


      '获取宽度和长度
    Dim Ld As ImageSize
    Ld = GetImageSize(Pathname)
        .SourceRect.Right = Ld.Width
        .SourceRect.Bottom = Ld.Height
        .Translation.x = 0
        .Translation.y = 0
        .Scale.x = 1
        .Scale.y = 1
        .Center.x = Int(Ld.Width / 2)
        .Center.y = Int(Ld.Height / 2)
        .Rotation = 0
        .Color = &HFFFFFFFF
        .Hnum = Hnum
        .Vnum = Vnum
        .Width = Ld.Width
        .Height = Ld.Height
        .MirrorType = xgNONE
        .FrmSpd = 100
        .CellIndex = 0
'        If .Hnum > 1 Or .Vnum > 1 Then
'            .CellIndex = 1
'        Else
'            .CellIndex = 0
'        End If
    End With
    
    
    
    Exit Function
ErrH:
    Debug.Print "Err [LoadGraph] 载入图片错误"
End Function
'绘图
Public Sub DrawGraph(ByVal x As Integer, ByVal y As Integer)
On Error GoTo ErrH:
    Dim ReturnValue As Long
    If gfx.D3DDevice Is Nothing Then End
    If gfx.D3DDevice.TestCooperativeLevel = D3DERR_DEVICELOST Then Exit Sub
    Sprite.Begin
            With gfx
                    .Translation.x = CSng(x) + (1 - .Scale.x) * gfx.Width / .Hnum / 2 '使缩放中心不变
                    .Translation.y = CSng(y) + (1 - .Scale.y) * gfx.Height / .Vnum / 2
                    If Not isMirrorOpen Then
                        Sprite.Draw .Texture, .SourceRect, .Scale, .Center, .Rotation, .Translation, .Color
                    Else
                        Dim matMirror As D3DMATRIX '做镜像用矩阵
                        Dim matTrans As D3DMATRIX '做平移矩阵
                        Dim matRotate As D3DMATRIX '做旋转矩阵
                        Dim matCenter As D3DMATRIX '将图形平移到原点
                        Dim matScale As D3DMATRIX
                        D3DXMatrixIdentity matMirror '单位矩阵
                        D3DXMatrixRotationZ matRotate, gfx.Rotation '旋转矩阵
                        D3DXMatrixScaling matScale, gfx.Scale.x, gfx.Scale.y, 1
                                   Select Case gfx.MirrorType
                            Case xgNONE
                                isMirrorOpen = False
                                D3DXMatrixIdentity MirrorMatrix
                            Case xgH_MIRROR
                                If .CellIndex <> 0 Then '精灵图情况
                                    D3DXMatrixTranslation matCenter, 0.5 * .Width \ .Hnum, -0.5 * .Height \ .Vnum, 0
                                    D3DXMatrixTranslation matTrans, x + 0.5 * .Width \ .Hnum, y + 0.5 * .Height \ .Vnum, 0
                                Else '正常图情况
                                    D3DXMatrixTranslation matCenter, gfx.Width \ 2, -gfx.Height \ 2, 0
                                    D3DXMatrixTranslation matTrans, x + gfx.Width \ 2, y + gfx.Height \ 2, 0
                                End If
                                
                                matMirror.m11 = -1
                                isMirrorOpen = True
                            Case xgV_MIRROR
                                If .CellIndex <> 0 Then '精灵图情况
                                    D3DXMatrixTranslation matCenter, -0.5 * gfx.Width \ .Hnum, 0.5 * gfx.Height \ .Vnum, 0
                                    D3DXMatrixTranslation matTrans, x + 0.5 * .Width \ .Hnum, y + 0.5 * .Height \ .Vnum, 0
                                Else '正常图情况
                                    D3DXMatrixTranslation matCenter, -gfx.Width \ 2, gfx.Height \ 2, 0
                                    D3DXMatrixTranslation matTrans, x + gfx.Width \ 2, y + gfx.Height \ 2, 0
                                End If
                                matMirror.m22 = -1
                                isMirrorOpen = True
                        End Select
                        
                        D3DXMatrixMultiply MirrorMatrix, matMirror, matCenter '输出一个已经平移到屏幕左上角镜像了的图像
                        D3DXMatrixMultiply MirrorMatrix, MirrorMatrix, matScale '移回到原始坐标
                        D3DXMatrixMultiply MirrorMatrix, MirrorMatrix, matRotate '在左上角旋转
                        D3DXMatrixMultiply MirrorMatrix, MirrorMatrix, matTrans '移回到原始坐标
                        
                        Sprite.DrawTransform .Texture, .SourceRect, MirrorMatrix, .Color '绘制图形
                        
                    End If
            End With
    Sprite.End
Exit Sub
ErrH:
    If gfx.Hnum = 1 And gfx.Vnum = 1 And gfx.CellIndex > 0 Then
        MsgBox "整图操作时禁止更改Cell值，这样是无意义的", vbCritical
        End
    Else
        Select Case Err.Number
            Case -2005530516
                MsgBox "[DrawGraph]错误的输入参数", vbCritical
                Debug.Print "[DrawGraph]错误的输入参数"
                Debug.Print "[DrawGraph]Cell= " & gfx.CellIndex
            Case 6 '溢出
                If gfx.Texture Is Nothing Then
                    MsgBox "[DrawGraph]没有可以显示的已载入图片", vbCritical
                ElseIf gfx.Hnum = 0 Or gfx.Vnum = 0 Then
                    MsgBox "[DrawGraph]载入图片时切分参数错误", vbCritical
                Else
                    MsgBox "[DrawGraph]参数错误,请检查输入参数"
                End If
            Case Else
                MsgBox "绘图时错误,请记下错误号并与sunicdavy@sina.com联系" & vbCr & "DX err:" & Err.Number, vbCritical
        End Select
        End
    End If
End Sub

'设置图像缩放比例
Public Sub SetScale(ByVal XScale As Single, ByVal YScale As Single)
    With gfx
        .Scale.x = XScale
        .Scale.y = YScale
        If .CellIndex = 0 Then
            .Center.x = XScale * gfx.SourceRect.Right / 2
            .Center.y = YScale * gfx.SourceRect.Bottom / 2
        Else
            .Center.x = .Width / .Hnum / 2 * XScale
            .Center.y = .Height / .Vnum / 2 * YScale
        End If
    End With
End Sub
'设置图像旋转角度
Public Sub SetRotate(ByVal Angle As Single)
    gfx.Rotation = Angle / 180 * 3.14
End Sub
'设置图像显示蒙板颜色
Public Sub SetColor(ByVal Color As ENUM_XG_COLOR)
  gfx.Color = Color
End Sub
'设置图像透明度
Public Sub SetAlpha(ByVal Degree As Integer)
    gfx.Color = D3DColorARGB(Degree, 255, 255, 255)
End Sub
'设置原图显示范围
Public Sub SetDisplayRect(ByVal Left As Single, ByVal Top As Single, ByVal Right As Single, ByVal Bottom As Single)
    With gfx.SourceRect
        .Left = Left
        .Top = Top
        .Right = Right
        .Bottom = Bottom
    End With
End Sub
'设置图形的镜像方式
Public Sub SetMirror(MirrorType As ENUM_XG_MIRROR)
    If MirrorType = xgH_MIRROR Or MirrorType = xgV_MIRROR Then
        isMirrorOpen = True
    Else
        isMirrorOpen = False
    End If
    gfx.MirrorType = MirrorType
End Sub
'获得图像宽度
Public Function Width() As Integer
    If gfx.CellIndex = 0 Then
        Width = gfx.Width
    Else
        Width = gfx.Width \ gfx.Hnum
    End If
End Function
'获得图像高度
Public Function Height() As Integer
    If gfx.CellIndex = 0 Then
        Height = gfx.Height
    Else
        Height = gfx.Height \ gfx.Vnum
    End If
End Function
'将D3D对象和设备载入到类中
Private Sub Class_Initialize()
    If D3DDevice Is Nothing Or D3DX Is Nothing Then
        Debug.Print "载入图片时没有初始化D3D设备和对象"
        Exit Sub
    End If
    Set gfx.D3DDevice = D3DDevice
    Set gfx.D3Dobj = D3DX
    Set Sprite = gfx.D3Dobj.CreateSprite(gfx.D3DDevice)
    isMirrorOpen = False
End Sub


Public Property Get Cell() As Integer
    Cell = gfx.CellIndex
End Property

Public Property Let Cell(ByVal Index As Integer)
On Error GoTo ErrH
    Dim eW As Integer
    Dim eH As Integer
    If Index < 0 Or Index > gfx.Hnum * gfx.Vnum Then
        'Debug.Print "[Cell]图片分割错误"
    End If
    With gfx
        eW = .Width \ .Hnum '单张图片的尺寸
        eH = .Height \ .Vnum
        If Index = 0 Then '正常图
           .CellIndex = 0
           .SourceRect.Left = 0
           .SourceRect.Top = 0
           .SourceRect.Right = .Width
           .SourceRect.Bottom = .Height
           .Center.x = Int(.Width / 2)
           .Center.y = Int(.Height / 2)
        Else '精灵图
            .CellIndex = Index
            .SourceRect.Left = eW * ((.CellIndex - 1) Mod .Hnum)
            .SourceRect.Top = eH * Int((.CellIndex - 1) / .Hnum)
            .SourceRect.Right = .SourceRect.Left + eW
            .SourceRect.Bottom = .SourceRect.Top + eH
            .Center.x = eW / 2
            .Center.y = eH / 2
        End If
    End With
Exit Property
ErrH:
Debug.Print "[Cell] 图片切换分块错误，请确认读取时的切分正确"
End

End Property
'读入文件的比例
Private Function GetImageSize(sFileName As String) As ImageSize
    On Error Resume Next
    Dim iFN As Integer
    Dim bTemp(3) As Byte
    Dim lFlen As Long
    Dim lPos As Long
    Dim bHmsb As Byte
    Dim bHlsb As Byte
    Dim bWmsb As Byte
    Dim bWlsb As Byte
    Dim bBuf(7) As Byte
    Dim bDone As Byte
    Dim iCount As Integer

    lFlen = FileLen(sFileName)
    iFN = FreeFile
    Open sFileName For Binary As iFN
    Get #iFN, 1, bTemp()
        
    'PNG 文件
    If bTemp(0) = &H89 And bTemp(1) = &H50 And bTemp(2) = &H4E _
    And bTemp(3) = &H47 Then
        Get #iFN, 19, bWmsb
        Get #iFN, 20, bWlsb
        Get #iFN, 23, bHmsb
        Get #iFN, 24, bHlsb
        GetImageSize.Width = CombineBytes(bWlsb, bWmsb)
        GetImageSize.Height = CombineBytes(bHlsb, bHmsb)
    End If
    
    'GIF 文件
    If bTemp(0) = &H47 And bTemp(1) = &H49 And bTemp(2) = &H46 _
    And bTemp(3) = &H38 Then
        Get #iFN, 7, bWlsb
        Get #iFN, 8, bWmsb
        Get #iFN, 9, bHlsb
        Get #iFN, 10, bHmsb
        GetImageSize.Width = CombineBytes(bWlsb, bWmsb)
        GetImageSize.Height = CombineBytes(bHlsb, bHmsb)
    End If
    
    
    'JPEG 文件
    If bTemp(0) = &HFF And bTemp(1) = &HD8 And bTemp(2) = &HFF Then
        lPos = 3
        Do
            Do
                Get #iFN, lPos, bBuf(1)
                Get #iFN, lPos + 1, bBuf(2)
                lPos = lPos + 1
            Loop Until (bBuf(1) = &HFF And bBuf(2) <> &HFF) Or lPos > lFlen
        
            For iCount = 0 To 7
                Get #iFN, lPos + iCount, bBuf(iCount)
            Next iCount
            If bBuf(0) >= &HC0 And bBuf(0) <= &HC3 Then
                bHmsb = bBuf(4)
                bHlsb = bBuf(5)
                bWmsb = bBuf(6)
                bWlsb = bBuf(7)
                bDone = 1
            Else
                lPos = lPos + (CombineBytes(bBuf(2), bBuf(1))) + 1
            End If
        Loop While lPos < lFlen And bDone = 0
        GetImageSize.Width = CombineBytes(bWlsb, bWmsb)
        GetImageSize.Height = CombineBytes(bHlsb, bHmsb)
        
    End If
    
    'BMP 文件
    If bTemp(0) = &H42 And bTemp(1) = &H4D Then
        Get #iFN, 19, bWlsb
        Get #iFN, 20, bWmsb
        Get #iFN, 23, bHlsb
        Get #iFN, 24, bHmsb
        GetImageSize.Width = CombineBytes(bWlsb, bWmsb)
        GetImageSize.Height = CombineBytes(bHlsb, bHmsb)
    End If
    
    
    'TGA 文件
    If bTemp(2) = &H2 And LCase(Right(sFileName, 4)) = ".tga" Then
        Get #iFN, 13, bWlsb
        Get #iFN, 14, bWmsb
        Get #iFN, 15, bHlsb
        Get #iFN, 16, bHmsb
        GetImageSize.Width = CombineBytes(bWlsb, bWmsb)
        GetImageSize.Height = CombineBytes(bHlsb, bHmsb)
    End If
    
    Close iFN
End Function
'getImageSize要用到的把十六进制数换成十进制
Public Function CombineBytes(lsb As Byte, msb As Byte) As Long
    CombineBytes = CLng(lsb + CLng(msb * 256))
End Function
'释放内存
Public Sub Release()
    With gfx
        Set .Texture = Nothing
        Set .D3DDevice = Nothing
        Set .D3Dobj = Nothing
        Set Sprite = Nothing
    End With
End Sub
'重新连接设备
Public Sub ReConnect()
    Set gfx.D3DDevice = D3DDevice
    Set gfx.D3Dobj = D3DX
    Set Sprite = gfx.D3Dobj.CreateSprite(gfx.D3DDevice)
End Sub
Public Property Get MaxFrame() As Integer
    MaxFrame = gfx.Hnum * gfx.Vnum
End Property

Public Property Get FrameSpeed() As Long
    FrameSpeed = gfx.FrmSpd
End Property

Public Property Let FrameSpeed(ByVal FrmSpd As Long)
    gfx.FrmSpd = FrmSpd
End Property

Public Sub GetColor(Alpha As Integer, Red As Integer, Green As Integer, Blue As Integer)
    Dim strColor As String
    strColor = Hex(gfx.Color)
    If Len(strColor) = 7 Then strColor = "0" & strColor
    Alpha = Hex2Dec(Left(strColor, 1)) * 16 + Hex2Dec(Mid(strColor, 2, 1))
    Red = Hex2Dec(Mid(strColor, 3, 1)) * 16 + Hex2Dec(Mid(strColor, 4, 1))
    Green = Hex2Dec(Mid(strColor, 5, 1)) * 16 + Hex2Dec(Mid(strColor, 6, 1))
    Blue = Hex2Dec(Mid(strColor, 7, 1)) * 16 + Hex2Dec(Mid(strColor, 8, 1))
End Sub
'将16进制字符转换为10进制数
Private Function Hex2Dec(HexStr As String) As Integer
    HexStr = LCase(HexStr)
    Select Case HexStr
        Case "a"
            Hex2Dec = 10
        Case "b"
            Hex2Dec = 11
        Case "c"
            Hex2Dec = 12
        Case "d"
            Hex2Dec = 13
        Case "e"
            Hex2Dec = 14
        Case "f"
            Hex2Dec = 15
        Case Else
            Hex2Dec = Val(HexStr)
    End Select
End Function

Public Function GetRotate() As Single
    GetRotate = gfx.Rotation
End Function

Public Sub GetScale(XScale As Single, YScale As Single)
    XScale = gfx.Scale.x
    YScale = gfx.Scale.y
    
End Sub

Public Function GetMirror() As ENUM_XG_MIRROR
    GetMirror = gfx.MirrorType
End Function

Public Function GetAlpha() As Integer
    Dim strColor As String
    strColor = Hex(gfx.Color)
    If Len(strColor) = 7 Then strColor = "0" & strColor
    GetAlpha = Hex2Dec(Left(strColor, 1)) * 16 + Hex2Dec(Mid(strColor, 2, 1))
End Function
'从资源文件里读取图片
Public Function LoadGraphFromRes(ByVal ResPathName As String, ByVal filename As String, ByVal ColorKey As ENUM_XG_COLOR, Optional Hnum As Integer = 1, Optional Vnum As Integer = 1) As Boolean
    On Error GoTo ErrH
    Dim Buf() As Byte
    Dim GfxInfo As D3DXIMAGE_INFO
    
    If Len(ResPathName) = 0 Then
        Debug.Print "Err: [LoadGraphFromRes] 路径不能为空"
        End
    Else
        '将默认文件名变为工作目录
        If Mid(ResPathName, 2, 1) <> ":" Then ResPathName = App.Path & "\" & ResPathName
        If Len(Dir(ResPathName)) = 0 Then
            Debug.Print "Err: [LoadGraphFromRes] 文件不存在"
            Debug.Print "     ResPathname:" & ResPathName
            End
        End If
    End If
    ResPathName = Trim(ResPathName) '去空格
    
    LoadGrf2Mem ResPathName, filename, Buf()
    
    With gfx
        Set .Texture = gfx.D3Dobj.CreateTextureFromFileInMemoryEx( _
                                                    .D3DDevice, _
                                                    Buf(0), _
                                                    UBound(Buf()), _
                                                    D3DX_DEFAULT, _
                                                    D3DX_DEFAULT, _
                                                    1, _
                                                    0, D3DFMT_UNKNOWN, _
                                                    D3DPOOL_MANAGED, _
                                                    D3DX_FILTER_NONE, _
                                                    D3DX_FILTER_NONE, _
                                                    ColorKey, ByVal 0, ByVal 0)
    Dim Ld As ImageSize
    Ld = GetImageSizeInMem(Buf())

      '获取宽度和长度
        .SourceRect.Right = Ld.Width
        .SourceRect.Bottom = Ld.Height
        .Translation.x = 0
        .Translation.y = 0
        .Scale.x = 1
        .Scale.y = 1
        .Center.x = Int(Ld.Width / 2)
        .Center.y = Int(Ld.Height / 2)
        .Rotation = 0
        .Color = &HFFFFFFFF
        .Hnum = Hnum
        .Vnum = Vnum
        .Width = Ld.Width
        .Height = Ld.Height
        .MirrorType = xgNONE
        .FrmSpd = 100
        If .Hnum > 1 Or .Vnum > 1 Then
            .CellIndex = 1
        Else
            .CellIndex = 0
        End If
    End With
    
    
    
    Exit Function
ErrH:
    Debug.Print "Err [LoadGraph] 载入图片错误"

End Function
'读取资源列表(和GetResFileInfo配合使用).返回文件头的长度
Private Sub LoadGrf2Mem(ByVal ResFilePath As String, ByVal filename As String, FileBuf() As Byte)
    Dim FileTag(2) As Byte
    Dim tag As Byte
    Dim TotalFileNum As Integer
    Dim tFileName As String
    Dim tFileSize As Long
    Dim tFileOffset As Long
    Dim ResBeginOffset As Long
    Dim ConvArr(1) As Byte
    Dim i As Integer
    Dim ResFileVer As Byte
    Dim ObjOffset As Long
    Dim ObjSize As Long
    tFileOffset = -1
    Open ResFilePath For Binary As 1
    'grf文件头确认
        Get 1, , FileTag(0)
        Get 1, , FileTag(1)
        Get 1, , FileTag(2)
        If Chr(FileTag(0)) <> "G" Or Chr(FileTag(1)) <> "R" Or Chr(FileTag(2)) <> "F" Then
            Debug.Print "This is not grf File"
            Close 1
            Exit Sub
        End If
        Get 1, , ResFileVer '版本号
        Get 1, , TotalFileNum '文件数
        For i = 1 To TotalFileNum
            Do
                DoEvents
                    Get 1, , tag '文件名称
                    If tag <> &HFF Then '若还没读完本文件记录
                            '连接中文字符
                            If Asc(Chr(tag)) = 0 Then '若是中文第一字符
                                ConvArr(0) = tag
                                Get 1, , tag
                                ConvArr(1) = tag
                                tFileName = tFileName & StrConv(ConvArr, vbUnicode)
                            Else
                                tFileName = tFileName & Chr(tag)
                            End If
                    Else
                         Exit Do
                    End If
            Loop
            Get 1, , tag
            Get 1, , tFileSize '文件大小
            Get 1, , tFileOffset '文件偏移
            If UCase(filename) = UCase(tFileName) Then
                ObjOffset = tFileOffset
                ObjSize = tFileSize
            Else
                tFileName = ""
            End If
        Next i
        
        If tFileOffset = -1 Then
                MsgBox "File not in Res"
                End
        End If
        Get 1, , tag 'FF标志
        If tag <> &HFF Then MsgBox "File Broke": End
        '寻址到指定文件位置
        Seek 1, Seek(1) + ObjOffset + 1
        ReDim FileBuf(ObjSize) As Byte
        '解开文件
        Get 1, , FileBuf()
        FileHeadMist FileBuf()
'        Open App.Path & "\1" & Filename For Binary As 2
'            Put 2, , FileBuf()
'        Close 2
    Close 1
    
End Sub
'每个文件内部头几百个字节的加密
Public Sub FileHeadMist(objBuf() As Byte)
    Dim i As Long
    Dim ThisMistLen As Long
    Const HeadMistLength As Integer = 250
    If UBound(objBuf()) < HeadMistLength Then
        ThisMistLen = UBound(objBuf())
    Else
        ThisMistLen = HeadMistLength
    End If
    For i = 0 To ThisMistLen
        objBuf(i) = objBuf(i) Xor 82
        objBuf(i) = objBuf(i) Xor 2
        objBuf(i) = objBuf(i) Xor 251
    Next i
End Sub
'从内存中的图片读取图片大小
Private Function GetImageSizeInMem(bTemp() As Byte) As ImageSize
    On Error Resume Next
    Dim iFN As Integer
    Dim lFlen As Long
    Dim lPos As Long
    Dim bHmsb As Byte
    Dim bHlsb As Byte
    Dim bWmsb As Byte
    Dim bWlsb As Byte
    Dim bBuf(7) As Byte
    Dim bDone As Byte
    Dim iCount As Integer
        
    'PNG 文件
    If bTemp(0) = &H89 And bTemp(1) = &H50 And bTemp(2) = &H4E _
    And bTemp(3) = &H47 Then
        bWmsb = bTemp(18)
        bWlsb = bTemp(19)
        bHmsb = bTemp(22)
        bHlsb = bTemp(23)
        GetImageSizeInMem.Width = CombineBytes(bWlsb, bWmsb)
        GetImageSizeInMem.Height = CombineBytes(bHlsb, bHmsb)
    End If
    
    'GIF 文件
    If bTemp(0) = &H47 And bTemp(1) = &H49 And bTemp(2) = &H46 _
    And bTemp(3) = &H38 Then
        bWlsb = bTemp(7)
        bWmsb = bTemp(8)
        bHlsb = bTemp(9)
        bHmsb = bTemp(10)
        GetImageSizeInMem.Width = CombineBytes(bWlsb, bWmsb)
        GetImageSizeInMem.Height = CombineBytes(bHlsb, bHmsb)
    End If
    
    
    'JPEG 文件
    If bTemp(0) = &HFF And bTemp(1) = &HD8 And bTemp(2) = &HFF Then
        lPos = 3
        Do
            Do
                bBuf(1) = bTemp(lPos)
                bBuf(2) = bTemp(lPos + 1)
                lPos = lPos + 1
            Loop Until (bBuf(1) = &HFF And bBuf(2) <> &HFF) Or lPos > UBound(bTemp())
        
            For iCount = 0 To 7
                bBuf(iCount) = bTemp(lPos + iCount)
            Next iCount
            If bBuf(0) >= &HC0 And bBuf(0) <= &HC3 Then
                bHmsb = bBuf(4)
                bHlsb = bBuf(5)
                bWmsb = bBuf(6)
                bWlsb = bBuf(7)
                bDone = 1
            Else
                lPos = lPos + (CombineBytes(bBuf(2), bBuf(1))) + 1
            End If
        Loop While lPos < UBound(bTemp()) And bDone = 0
        GetImageSizeInMem.Width = CombineBytes(bWlsb, bWmsb)
        GetImageSizeInMem.Height = CombineBytes(bHlsb, bHmsb)
        
    End If
    
    'BMP 文件
    If bTemp(0) = &H42 And bTemp(1) = &H4D Then
        bWlsb = bTemp(18)
        bWmsb = bTemp(19)
        bHlsb = bTemp(22)
        bHmsb = bTemp(23)
        GetImageSizeInMem.Width = CombineBytes(bWlsb, bWmsb)
        GetImageSizeInMem.Height = CombineBytes(bHlsb, bHmsb)
    End If
    
    
    'TGA 文件
    If bTemp(2) = &H2 Then
        bWlsb = bTemp(12)
        bWmsb = bTemp(13)
        bHlsb = bTemp(14)
        bHmsb = bTemp(15)
        GetImageSizeInMem.Width = CombineBytes(bWlsb, bWmsb)
        GetImageSizeInMem.Height = CombineBytes(bHlsb, bHmsb)
    End If

End Function
