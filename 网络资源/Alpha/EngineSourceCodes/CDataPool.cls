VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
'**模 块 名：CDataPool
'**说    明：Boywhp 版权所有2005 - 2006(C)
'**创 建 人：王慧平
'**日    期：2004-11
'**修 改 人：王慧平
'**日    期：2005-07-05 15:52:47
'**描    述：高效率的动态内存分配，将数据进行二进制打包以及解包
'**版    本：V1.0.0
'*************************************************************************
Option Explicit

Private Const MIN_MEM_SIZE = 65535 ' &HFFFF

Private m_Datas() As Byte
Private m_Offset As Long                '偏移
Private m_lpData As Long                '源数据地址
Private m_DataLenth As Long
Private m_DisableBuffer As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Public Sub GetBytes(DataBytes() As Byte)
    DataBytes = m_Datas
End Sub

Private Function Malloc(size As Long) As Long
    '自动切割数据适应 MIN_MEM_SIZE
    Dim freeSize As Long, tSize As Long
    If m_DisableBuffer Then
        tSize = size
    Else
        freeSize = UBound(m_Datas) - m_Offset
        tSize = size - freeSize             '实际应该分配的空间
        If tSize > 0 Then
            tSize = (tSize And &HFFFF0000) + MIN_MEM_SIZE
        End If
    End If
    
    If tSize > 0 Then
        ReDim Preserve m_Datas(UBound(m_Datas) + tSize)
        m_lpData = VarPtr(m_Datas(0))
    End If
    
    Malloc = -1
End Function

Public Sub WriteBytes(lpsrc As Long, size As Long, Optional appended As Boolean = True)
    '默认
    If appended Then
        m_Offset = m_DataLenth
        Malloc size
        CopyMemory ByVal (m_lpData + m_Offset), ByVal lpsrc, size
        m_Offset = m_Offset + size
    
        If m_Offset > m_DataLenth Then
            m_DataLenth = m_Offset
        End If
    Else
        ModifyData m_Offset, lpsrc, size
    End If
End Sub

Public Function ModifyData(offset As Long, lpsrc As Long, size As Long) As Long
    '不修改偏移地址
    Debug.Assert offset + size < m_DataLenth
    CopyMemory ByVal (m_lpData + offset), ByVal lpsrc, size
End Function

Public Function GetLenth() As Long
    GetLenth = m_Offset
End Function

Public Sub Clear()
    ReDim m_Datas(0)
    m_Offset = 0
    m_DataLenth = 0
    m_lpData = VarPtr(m_Datas(0))
End Sub

Public Sub SetUpDataPointer(datas() As Byte)
    m_Offset = 0
    m_lpData = VarPtr(datas(0))
    m_DataLenth = UBound(datas) + 1
End Sub

Public Sub ReSetPointer(lpDatas As Long, size As Long)
    '设置数据区指针
    m_Offset = 0
    m_lpData = lpDatas
    m_DataLenth = size
End Sub

Public Sub MovePointer(offset As Long)
    Debug.Assert offset >= 0 And offset < m_DataLenth
    m_Offset = offset
End Sub

Public Sub CopyDatas(datas() As Byte)
    ReDim datas(m_DataLenth - 1)
    CopyMemory datas(0), m_Datas(0), m_DataLenth    'UBound(m_Datas)+1 ???
End Sub

Public Sub AppendDatas(datas() As Byte)
    '追加数据
    WriteBytes VarPtr(datas(0)), UBound(datas) + 1
End Sub

Public Sub AddLong(ByVal data As Long)
    '32bits
    WriteBytes VarPtr(data), 4
End Sub

Public Sub AddFloat(ByVal data As Single)
    '32bits
    WriteBytes VarPtr(data), 4
End Sub

Public Sub AddInteger(ByVal data As Integer)
    '16bits
    WriteBytes VarPtr(data), 2
End Sub

Public Sub AddString(ByVal s As String)
    '添加数据,vbFromUnicode
    Dim slen As Long
    AddLong CLng(Len(s))
    
    s = StrConv(s, vbFromUnicode)
    slen = LenB(s)
    
    AddLong slen
    WriteBytes StrPtr(s), slen
End Sub

Public Sub AddDefType(ByVal lpObj As Long, ByVal objLenb As Long)
    Debug.Assert objLenb > 0 And lpObj <> 0
    AddLong objLenb
    WriteBytes lpObj, objLenb
End Sub

Public Function GetLong(Optional ByVal offset As Long = -1) As Long
    Debug.Assert offset >= 0 Or offset = -1
    
    '-1 模式，自动移位
    If offset = -1 Then
        CopyMemory GetLong, ByVal (m_lpData + m_Offset), 4
        m_Offset = m_Offset + 4
    ElseIf offset >= 0 Then
        CopyMemory GetLong, ByVal (m_lpData + offset), 4
        m_Offset = offset + 4
    End If
End Function

Public Function GetFloat(Optional ByVal offset As Long = -1) As Single
    Debug.Assert offset >= 0 Or offset = -1
    '-1 模式，自动移位
    If offset = -1 Then
        CopyMemory GetFloat, ByVal (m_lpData + m_Offset), 4
        m_Offset = m_Offset + 4
    ElseIf offset >= 0 Then
        CopyMemory GetFloat, ByVal (m_lpData + offset), 4
        m_Offset = offset + 4
    End If
End Function

Public Function GetInteger(Optional ByVal offset As Long = -1) As Integer
    Debug.Assert offset >= 0 Or offset = -1
    
    '-1 模式，自动移位
    If offset = -1 Then
        CopyMemory GetInteger, ByVal (m_lpData + m_Offset), 2
        m_Offset = m_Offset + 4
    ElseIf offset >= 0 Then
        CopyMemory GetInteger, ByVal (m_lpData + offset), 2
        m_Offset = offset + 4
    End If
End Function

Public Function GetString(Optional ByVal offset As Long = -1) As String
    Dim slen As Long, s As String, stlen As Long
    stlen = GetLong(offset)
    slen = GetLong(offset)
    
    s = Space(slen)
    CopyMemory ByVal StrPtr(s), ByVal (m_lpData + m_Offset), slen
    s = Left(StrConv(s, vbUnicode), stlen)
    m_Offset = m_Offset + slen
    GetString = s
End Function

Public Sub GetDefType(ByVal lpObj As Long, Optional ByVal offset As Long = -1)
    Dim objlen As Long
    objlen = GetLong(offset)
    
    Debug.Assert objlen > 0 And lpObj <> 0
    '使用GetLong后自动处理好了移位了
    CopyMemory ByVal lpObj, ByVal (m_lpData + m_Offset), objlen
    m_Offset = m_Offset + objlen
End Sub

Public Sub GetDatas(ByVal lpDst As Long, size As Long, Optional ByVal offset As Long = -1)
    If offset = -1 Then
        CopyMemory ByVal lpDst, ByVal (m_lpData + m_Offset), size
        m_Offset = m_Offset + size
    ElseIf offset >= 0 Then
        CopyMemory ByVal lpDst, ByVal (m_lpData + offset), size
        m_Offset = offset + size
    End If
End Sub

Private Sub Class_Initialize()
    Clear
End Sub


