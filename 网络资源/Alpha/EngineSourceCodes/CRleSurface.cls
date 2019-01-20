VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRleSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Rle压缩解压代码
Private Type TFileHeader
    Bytes As Long
    ID As Long
    Style As Long                   '类型
    Reserve As Long                 '保留
End Type

Private Type TRle_Headerinfo
    Bytes As Long                  '起点数
    ID As Long
    W As Long
    H As Long
    x As Long
    y As Long
    pitch As Long
End Type

Private m_RleBuffer As CDataPool
Private m_Rle_Header As TRle_Headerinfo
Private m_ID As Long
Private m_lpRleDatas As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function RGB565 Lib "alpha.dll" (ByVal RGB555 As Integer) As Integer

Public Function GetlpRleDatas() As Long
    GetlpRleDatas = m_lpRleDatas
End Function

Public Function GetRlePointsNum() As Long
    GetRlePointsNum = m_Rle_Header.Bytes
End Function

Public Sub LoadFile(lpRleFile As Long)
    '载入数据从数据包
    Set m_RleBuffer = Nothing
    Set m_RleBuffer = New CDataPool

    '设置数据区指针 20为通用文件头 objFileheaderSize=16+4
    Dim tmpHeader As TFileHeader
    m_RleBuffer.ReSetPointer lpRleFile, 0
    m_RleBuffer.GetDefType VarPtr(tmpHeader)
    m_RleBuffer.GetDefType VarPtr(m_Rle_Header)

    Dim flags As Long
    flags = m_RleBuffer.GetLong
    Debug.Assert flags = -1

    m_lpRleDatas = lpRleFile + m_RleBuffer.GetLenth
    'Debug.Print m_RleBuffer.GetLenth
    '注意有一个文件标示FFFF 4
    m_RleBuffer.ReSetPointer m_lpRleDatas, tmpHeader.Bytes - LenB(m_Rle_Header) - 8

End Sub

Public Function GetBmpSurface(ID As Long) As CGDISurface
    '解码Suface
    Dim offset As Long
    Dim i As Long, j As Long

    '快速定位
    '开始解码
    m_RleBuffer.MovePointer 0

    Set GetBmpSurface = New CGDISurface
    GetBmpSurface.Create m_Rle_Header.W, m_Rle_Header.H

    Dim colorDatas() As Byte, x As Long, y As Long, lenth As Long
    Dim dOffset As Long

    GetBmpSurface.LoadData colorDatas
    For j = 0 To m_Rle_Header.Bytes - 1
        x = m_RleBuffer.GetLong
        y = m_RleBuffer.GetLong
        lenth = m_RleBuffer.GetLong

        dOffset = dOffset + x + y * m_Rle_Header.pitch
        m_RleBuffer.GetDatas VarPtr(colorDatas(dOffset)), lenth * 2
    Next
    GetBmpSurface.SetData colorDatas
    Set m_RleBuffer = Nothing

Exit Function

    Set GetBmpSurface = Nothing
End Function

Public Sub Blt(srcPic As CGDISurface, frameID As Long, _
               Optional x As Long = 0, _
               Optional y As Long = 0, _
               Optional RGB555 As Boolean = True, _
               Optional appended As Boolean = False)
    '进行Rle压缩
    Dim srcData() As Byte, dataSize As Long, oldLen As Long
    Dim rleHeader As TRle_Headerinfo

    rleHeader.W = srcPic.Width
    rleHeader.H = srcPic.Height
    rleHeader.pitch = srcPic.pitch
    rleHeader.x = x
    rleHeader.y = y
    rleHeader.ID = frameID

    If Not appended Then
        Set m_RleBuffer = Nothing
        Set m_RleBuffer = New CDataPool
    End If

    oldLen = m_RleBuffer.GetLenth

    srcPic.LoadData srcData, 1
    RLEScan rleHeader, srcData, RGB555

    '修改headerinfo
    dataSize = m_RleBuffer.GetLenth - oldLen ' (UBound(dstData) + 1)

End Sub

Public Function GetRleData() As CDataPool
    Set GetRleData = m_RleBuffer
End Function

Private Sub RLEScan(rleHeader As TRle_Headerinfo, srcData() As Byte, Optional RGB555 As Boolean = True)
    '16bits RGB555 透明色压缩
    Dim i As Long
    Dim color As Long, oldColor As Long
    Dim x As Long, y As Long, oldx As Long, oldy As Long
    Dim oldPos As Long, PointNum As Long, colorNum As Long, headerPos As Long

    headerPos = m_RleBuffer.GetLenth + 4                        '注意前四个字节保留
    m_RleBuffer.AddDefType VarPtr(rleHeader), LenB(rleHeader)
    m_RleBuffer.AddLong &HFFFFFFFF

    oldColor = &H7C1F
    colorNum = 0
    PointNum = 0
    For i = 0 To UBound(srcData) Step 2
        CopyMemory color, srcData(i), 2
        If color = &H7C1F Then
            If oldColor <> &H7C1F Then
                '记录结束点Lenth
                m_RleBuffer.ModifyData oldPos + 8, VarPtr(colorNum), 4
                colorNum = 0
            End If
        ElseIf i Mod rleHeader.pitch < 2 * rleHeader.W Then
            If oldColor = &H7C1F Then
                '记录起点(x,y)
                oldPos = m_RleBuffer.GetLenth

                x = i Mod rleHeader.pitch
                y = i / rleHeader.pitch - 0.4999
                
                'Debug.Print x, y
                
                m_RleBuffer.AddLong x - oldx
                m_RleBuffer.AddLong y - oldy
                m_RleBuffer.AddLong 0

                oldx = x
                oldy = y

                colorNum = 0
                PointNum = PointNum + 1
            End If

            If RGB555 Then
                m_RleBuffer.AddInteger color
            Else
                m_RleBuffer.AddInteger RGB565(color)
            End If
            colorNum = colorNum + 1

        End If
        oldColor = color

        If i Mod rleHeader.pitch = rleHeader.pitch - 2 Then
            '结尾强制添加
            oldColor = &H7C1F
            If colorNum > 0 Then
                m_RleBuffer.ModifyData oldPos + 8, VarPtr(colorNum), 4
                colorNum = 0
            End If
        End If
    Next
    'Rle OK!
    Debug.Assert colorNum = 0
    
    m_RleBuffer.ModifyData headerPos, VarPtr(PointNum), 4

End Sub

Public Sub RenderTo(x As Long, y As Long, dstSurface As CGDISurface)
    rle_blt ByVal dstSurface.lpData, dstSurface.pitch, dstSurface.Height, x, y, _
            ByVal Me.GetlpRleDatas, Me.GetRlePointsNum
End Sub
