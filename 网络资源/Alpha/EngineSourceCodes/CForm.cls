VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Left As Integer
Public Top As Integer
Public Right As Integer
Public Bottom As Integer
Public Actived As Boolean
Public Visiable As Boolean
Public ModleShow As Boolean                     '是否独占直到关闭
Public Caption  As String
Public BackPic As CSurface
Public Frame As Integer
Public WithEvents MsgEngine As CGameEngine      '需要引擎的消息支持
Private m_HotRects As CHotRect                  '区域树

Public Sub GetMessages(MsgType As InputMsgType, ByVal Value As Long)
    '处理消息
    Caption = Str(Value)
End Sub

Public Sub Render()
    'Test
    
End Sub

Public Function CheckRect(ByVal X As Integer, ByVal Y As Integer) As Boolean
    If X > Left And X < Right And Y > Top And Y < Bottom Then CheckRect = True
End Function

Private Sub MsgEngine_InputMsg(MsgType As InputMsgType, ByVal Value As Long, ByVal X As Long, ByVal Y As Long)
    m_HotRects.TreatMsg MsgType, Value, X, Y
End Sub
