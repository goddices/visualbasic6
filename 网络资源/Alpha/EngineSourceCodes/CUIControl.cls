VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUIControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'控件接口
Public Event OnPaint()
Public Event OnInputMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)
Public Event OnCheckRect(ret As Boolean, ByVal x As Long, ByVal y As Long)
Public Event OnRectAway()       '区域离开事件，暂时不支持！

Public ID As Long
Public Left As Long
Public Top As Long
Public Width As Long
Public Height As Long
Public Visiable As Boolean
Public Release As Boolean
Public Parent As CUIForm
Public Actived As Boolean

Public Sub SendInputMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)
    RaiseEvent OnInputMsg(MsgType, Value, x, y)
End Sub

Public Sub Render()
    RaiseEvent OnPaint
End Sub

Public Function IsInRect(ByVal x As Long, ByVal y As Long) As Boolean
    RaiseEvent OnCheckRect(IsInRect, x, y)
End Function

Public Sub Create(x As Integer, y As Integer, w As Integer, h As Integer)
    Debug.Assert w > 0 And h > 0
    Debug.Assert Not Parent Is Nothing
    Width = w
    Height = h
    Left = x
    Top = y

    Release = False
    Parent.RegControl Me
End Sub

