VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBackMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------处理地图的滚动显示--------
'Buffer为超宽缓冲图，绘制mViewRect到g_MainSurface
'Create dw,dh,sw,sh:
'RefillBuffer
'SetViewRect x,y:移动视窗左上-》x,y
'Render

Public Buffer As CSurface

Private mViewRect As RECT           '相对于Buffer的REct
Private mViewWidth As Integer
Private mViewHeight As Integer
Private mMapWidth As Integer
Private mMapHeight As Integer
Private mLeft As Integer            '整个Buffer相对于地图的坐标
Private mTop As Integer

Public Event RefillBuffer(ByVal x As Integer, ByVal y As Integer)          '请求刷新缓冲

Private Sub Class_Terminate()
    Set Buffer = Nothing
End Sub


'整个可视区的原地图坐标
Public Property Get MapX() As Integer
    MapX = mLeft + mViewRect.Left
End Property

Public Property Get MapY() As Integer
    MapY = mTop + mViewRect.Top
End Property

Private Sub RefreshRect(ByVal x As Integer, ByVal y As Integer)
    '刷新可视区
    mViewWidth = g_MainSurface.GetClipper.Right - g_MainSurface.GetClipper.Left
    mViewHeight = g_MainSurface.GetClipper.Bottom - g_MainSurface.GetClipper.Top
    mViewRect.Left = x - mLeft
    mViewRect.Top = y - mTop
    mViewRect.Right = mViewRect.Left + mViewWidth
    mViewRect.Bottom = mViewRect.Top + mViewHeight
End Sub

Private Sub DefaultViewRect()
    SetViewRect mLeft + (Buffer.width - mViewWidth) / 2, mTop + (Buffer.height - mViewHeight) / 2
End Sub

Public Sub SetBufferRect(x As Integer, y As Integer)
    mLeft = x
    mTop = y
    RaiseEvent RefillBuffer(mLeft, mTop)
End Sub

Public Sub SetViewRect(x As Integer, y As Integer)
    Dim OutBound As Boolean
    'x,y为地图的坐标
    If x < (Buffer.width - mViewWidth) / 2 Then x = (Buffer.width - mViewWidth) / 2
    If y < (Buffer.height - mViewHeight) / 2 Then y = (Buffer.height - mViewHeight) / 2
    If x > mMapWidth - mViewWidth Then x = mMapWidth - mViewWidth
    If y > mMapHeight - mViewHeight Then y = mMapHeight - mViewHeight
    '防止越界
    RefreshRect x, y
    If mViewRect.Left < 0 Then
        mLeft = mLeft - (Buffer.width - mViewWidth) / 2
        OutBound = True
    End If
    
    If mViewRect.Top < 0 Then
        mTop = mTop - (Buffer.height - mViewHeight) / 2
        OutBound = True
    End If
    
    If mViewRect.Right > Buffer.width Then
        mLeft = mLeft + (Buffer.width - mViewWidth) / 2
        OutBound = True
    End If
    
    If mViewRect.Bottom > Buffer.height Then
        mTop = mTop + (Buffer.height - mViewHeight) / 2
        OutBound = True
    End If
    
    If OutBound Then
        RaiseEvent RefillBuffer(mLeft, mTop)
        RefreshRect x, y    '注意要RefreshRect,因为mLeft,mtop已经改变
    End If
End Sub

Public Function Create(ByVal dw As Integer, ByVal dh As Integer, ByVal sw As Integer, ByVal sh As Integer) As Boolean
    'dw,dh缓冲区，sw,sh 地图大小
    On Error GoTo errhand
    
    mViewWidth = g_ViewRect.Right
    mViewHeight = g_ViewRect.Bottom
    If dw < mViewWidth Or dh < mViewHeight Then GoTo errhand
    
    Set Buffer = New CSurface
    Buffer.Create dw, dh
    
    mMapWidth = sw
    mMapHeight = sh
    
    DefaultViewRect
    Create = True
    Exit Function
    '----
errhand:

End Function

Public Sub Render()
    g_MainSurface.RectBlt Buffer, g_MainSurface.GetClipper, mViewRect
End Sub

Public Sub GetViewPoint(x As Integer, y As Integer)
    '得到地图相对于视窗的坐标
    x = x - mLeft - mViewRect.Left
    y = y - mTop - mViewRect.Top
End Sub

Public Sub GetBufferPoint(x As Integer, y As Integer)
    '得到地图相对于Buffer的坐标
    x = x - mLeft
    y = y - mTop
End Sub

Public Function GetViewRect() As RECT
    '返回View相对Buffer的REct
    GetViewRect = mViewRect
End Function
