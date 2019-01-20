VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDemoMsgPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---------------- Class MessagePost --------------
'             处理全局的消息邮递
'-------------------------------------------------
Private Const MAX_MSG_NUM = 100
Private m_MsgStack(MAX_MSG_NUM) As String                   '消息分配
Private m_StackTopPointer As Integer

Public Sub PostMessage(Msg As String)
    m_MsgStack(m_StackTopPointer) = Msg
    m_StackTopPointer = m_StackTopPointer + 1
    If m_StackTopPointer > MAX_MSG_NUM Then m_StackTopPointer = 0
End Sub

Private Sub Class_Initialize()
    m_StackTopPointer = 0
End Sub

Public Function PeekMessage(Msg As String) As Integer
    '查找消息
    Dim i As Integer
    For i = m_StackTopPointer - 1 To 0 Step -1
        If m_MsgStack(i) = Msg Then PeekMessage = i: Exit Function
    Next i
    PeekMessage = -1
End Function

Public Function GetMessage(MsgAddress As Integer, Optional DelMsg As Boolean = True) As String
    Dim i As Integer
    If MsgAddress < 0 Or MsgAddress > m_StackTopPointer Then Exit Function
    GetMessage = m_MsgStack(MsgAddress)
    '取出消息后删除消息
    If DelMsg Then
        For i = MsgAddress To m_StackTopPointer - 2
            m_MsgStack(i) = m_MsgStack(i + 1)
        Next i
        'Del the msg
        m_StackTopPointer = m_StackTopPointer - 1
        If m_StackTopPointer < 0 Then m_StackTopPointer = 0: g_Err_Description = "消息栈越出!"
    End If
End Function

