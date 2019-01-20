VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilePacket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'压缩包打包解包代码 by boywhp
Private Type TPacketHeader
    PacketFlags As Long
    PacketID As Long
    
    FileNum As Long
    FileDataSize As Long
    ChildNum As Long
    
    lp_File_Table As Long
    lp_File_Base As Long
    
    lp_Child_Table As Long
    lp_Child_Base As Long
End Type

Private Type TFile_Tableinfo
    FileFrameID As Long
    lp_FileDatas As Long
End Type

Private Type TChild_Tableinfo
    ChildID As Long
    lp_Child As Long
End Type

Private Type TFileHeader
    Bytes As Long
    ID As Long
    Style As Long                   '文件类型
    Reserve As Long                 '保留
End Type

Private m_lpDataBase As Long
Private m_Header As TPacketHeader
Private m_DataLenth As Long
Private m_Buffers() As Byte         'File_Datas
Private m_BufferPool As CDataPool   '便于操作数据
Private m_FileNum As Long

Private m_File_Tables() As TFile_Tableinfo
Private m_Childs_Tables() As TChild_Tableinfo
Private m_Childs() As CFilePacket     '子包

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Public Sub Save()
    Dim i As Long
    'Header
    Put #m_FileNum, , m_Header
    'FileTables
    For i = 0 To m_Header.FileNum - 1
        Put #m_FileNum, , m_File_Tables(i)
    Next
    'FileDatas
    Debug.Assert Not m_BufferPool Is Nothing
    m_BufferPool.CopyDatas m_Buffers
    Put #1, , m_Buffers
    'childTables
    For i = 0 To m_Header.ChildNum - 1
        Put #m_FileNum, , m_Childs_Tables(i)
    Next
    'Child Datas
    '。。。
    Close #m_FileNum
End Sub

Public Sub AppendFile(FileName As String, datas() As Byte)
    'Append File
    '修改headerinfo
    Dim dataSize As Long
    Dim tmpFileHeader As TFileHeader
    Dim oldLenth As Long
    
    If m_BufferPool Is Nothing Then
        Set m_BufferPool = New CDataPool
    End If
    
    tmpFileHeader.Bytes = UBound(datas) + 1
    tmpFileHeader.ID = Hash(FileName)
    tmpFileHeader.Style = 0
    tmpFileHeader.Reserve = 0
    
    oldLenth = m_BufferPool.GetLenth
    m_BufferPool.AddDefType VarPtr(tmpFileHeader), LenB(tmpFileHeader)
    m_BufferPool.AppendDatas datas
    
    dataSize = m_BufferPool.GetLenth - oldLenth
    
    With m_Header
        .FileNum = m_Header.FileNum + 1
        .FileDataSize = .FileDataSize + dataSize
        
        ReDim Preserve m_File_Tables(.FileNum - 1)
        m_File_Tables(.FileNum - 1).FileFrameID = tmpFileHeader.ID
        m_File_Tables(.FileNum - 1).lp_FileDatas = .lp_Child_Table
        
        Dim i As Long
        For i = 0 To .FileNum - 1
            m_File_Tables(i).lp_FileDatas = m_File_Tables(i).lp_FileDatas + LenB(m_File_Tables(0))
        Next
        
        .lp_File_Table = LenB(m_Header)
        .lp_File_Base = .lp_File_Base + LenB(m_File_Tables(0))
        .lp_Child_Base = .lp_Child_Base + dataSize + LenB(m_File_Tables(0))
        .lp_Child_Table = .lp_Child_Table + dataSize + LenB(m_File_Tables(0))
    End With
End Sub

Public Sub OpenPacketFile(FileName As String, Optional appended As Boolean = False)
    m_FileNum = FreeFile
    Open FileName For Binary As #m_FileNum
    If Not appended Then
        '设置文件头
        m_Header.PacketFlags = &H9910
        m_Header.PacketID = &H177
        
        m_Header.FileNum = 0
        m_Header.FileDataSize = 0
        m_Header.lp_File_Table = LenB(m_Header)
        m_Header.lp_File_Base = LenB(m_Header)
        
        m_Header.ChildNum = 0
        m_Header.lp_Child_Table = m_Header.lp_File_Base + m_Header.FileDataSize
        m_Header.lp_Child_Base = m_Header.lp_File_Base + m_Header.FileDataSize
    End If
End Sub

Public Sub ReadPacketFile(FileName As String)
    '载入数据包
    Dim i As Long
    m_FileNum = FreeFile
    Open FileName For Binary As #m_FileNum
    Get #m_FileNum, , m_Header
    
    ReDim m_File_Tables(m_Header.FileNum - 1)
    For i = 0 To m_Header.FileNum - 1
        Get #m_FileNum, , m_File_Tables(i)
    Next
    'GetFileDatas()
    ReDim m_Buffers(m_Header.FileDataSize - 1)
    Get #m_FileNum, , m_Buffers
    'Get Child Table
    If m_Header.ChildNum > 0 Then
        ReDim m_Childs_Tables(m_Header.ChildNum - 1)
    End If
    For i = 0 To m_Header.ChildNum - 1
        Get #m_FileNum, , m_Childs_Tables(i)
    Next
    'Get Child Datas
    '。。。
End Sub

Public Function GetOffsetBase() As Long
    GetOffsetBase = m_lpDataBase
End Function

Public Sub ResetFileBase(FileBase As Long)
    m_lpDataBase = FileBase
End Sub

Public Function GetPacketLenth() As Long
    GetPacketLenth = m_DataLenth
End Function

Public Function GetFileNum() As Long
    GetFileNum = m_Header.FileNum
End Function

Public Function GetFile(FileName As String, Optional index As Long = -1, Optional ID As Long = -1) As Long
    Dim i As Long, offset As Long
    '返回内存映射的文件首地址的指针
    If index >= 0 Then
        offset = m_File_Tables(index).lp_FileDatas - m_Header.lp_File_Base
        GetFile = VarPtr(m_Buffers(offset))
        Exit Function
    End If
    
    ID = IIf(ID = -1, Hash(FileName), ID)
    
    For i = 0 To GetFileNum - 1
        If m_File_Tables(i).FileFrameID = ID Then
            offset = m_File_Tables(i).lp_FileDatas - m_Header.lp_File_Base
            GetFile = VarPtr(m_Buffers(offset))
            '注意：这里返回的是真实的文件头跟用户直接添加的有出入
            '真实的数据开始为GetFile+20
            Exit Function
        End If
    Next
    GetFile = -1
End Function

Public Function Hash(s As String) As Long
    Dim i As Long
    's请使用英文字符
    Dim size As Long
    size = &H9910
    s = UCase(s)
    For i = 1 To Len(s)
        Hash = (Hash * 26 + (Asc(Mid(s, i, 1)) - 65)) Mod size
    Next
End Function

Public Sub AppendChild(child As CFilePacket)
    '暂时不支持
End Sub
Public Function GetChild(ChildID As Long) As CFilePacket
    '暂时不支持
End Function
