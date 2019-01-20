VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------Class INI File----------------------------
'                   提供ini文件的读写操作
'------------------------------------------------------------------

Private m_FileName As String

Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" _
    (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" _
    (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
    
Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal NewFile As String)
    m_FileName = NewFile
End Property

Public Sub WriteLineString(ByVal AppName As String, ByVal s As String)
    Dim i As Long
    i = WritePrivateProfileSection(AppName, s, m_FileName)
    '注意:当写入下一行时 若AppName相同则以前写入的将被清除！
End Sub

Public Sub ClearSection(ByVal AppName As String)
    Dim i As Long
    i = WritePrivateProfileSection(AppName, vbNullString, m_FileName)
End Sub

Public Sub WriteString(ByVal AppName As String, ByVal KeyName As String, ByVal Value As String)
    Dim i As Long
    i = WritePrivateProfileString(AppName, KeyName, Value, m_FileName)
    '注意:当写入下一行时 若AppName相同则以前写入的将被清除！
End Sub

Public Function GetSectionString(ByVal AppName As String, ByVal KeyName As String) As String
    Dim s As String * 255, slength As Long
    slength = GetPrivateProfileString(AppName, KeyName, "Error", s, 255, m_FileName)
    '注意使用的是BYTES,LeftB
    slength = slength - (LenB(StrConv(s, vbFromUnicode)) - 255)
    If slength > 0 Then GetSectionString = Left(s, slength)
End Function

Public Function GetSectionInt(ByVal AppName As String, ByVal KeyName As String, Optional Default As Long = 0) As Long
    GetSectionInt = GetPrivateProfileInt(AppName, KeyName, Default, m_FileName)
End Function

Public Function GetSection(ByVal AppName As String) As String
    Dim s As String * 255, slength As Long
    '得到指定项目下的所有项。若缓冲区溢出则返回 255－2
    slength = GetPrivateProfileSection(AppName, s, 255, m_FileName)
    GetSection = Left(s, slength)
End Function
