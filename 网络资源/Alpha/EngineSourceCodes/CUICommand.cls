VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUICommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Caption As String
Public WithEvents UIControl As CUIControl       '控件接口
Attribute UIControl.VB_VarHelpID = -1

Private m_State As Long
Private m_MouseMoving As Boolean

Private Sub Class_Initialize()
    Set UIControl = New CUIControl
End Sub

Private Sub UIControl_OnCheckRect(ret As Boolean, ByVal x As Long, ByVal y As Long)
    Dim tmpRect As RECT
    tmpRect = UIControl.Parent.GetWindowRect
    tmpRect.Left = UIControl.Left + tmpRect.Left
    tmpRect.Top = UIControl.Top + tmpRect.Top
    tmpRect.Right = tmpRect.Left + UIControl.Width
    tmpRect.Bottom = tmpRect.Top + UIControl.Height
    ret = CheckRect(tmpRect, x, y)
End Sub

Private Sub UIControl_OnInputMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)
    '处理消息
    Caption = Str(MsgType) + Str(Value) + Str(x)
     
    If MsgType = MSG_MouseMove Then
        m_MouseMoving = True
        UIControl.Release = True
    End If
    
    Dim bSelect As Boolean
    Call UIControl_OnCheckRect(bSelect, x, y)
    UIControl.Release = bSelect
    m_MouseMoving = bSelect
    
End Sub

Private Sub UIControl_OnPaint()
    '处理绘制代码
    Dim tmpRect As RECT
    tmpRect = UIControl.Parent.GetWindowRect
    
    With g_MainSurface.DD_Surface
        If m_MouseMoving Then
            .SetForeColor vbBlue
            
        Else
            .SetForeColor vbWhite
        End If
        .SetFillColor RGB(220, 100, 100)
        .DrawRoundedBox UIControl.Left + tmpRect.Left, _
                        UIControl.Top + tmpRect.Top, _
                        UIControl.Left + tmpRect.Left + UIControl.Width, _
                        UIControl.Top + tmpRect.Top + UIControl.Height, _
                        10, 10
        If UIControl.Actived Then
            .setDrawStyle 2
            .DrawRoundedBox UIControl.Left + tmpRect.Left + 2, _
                            UIControl.Top + tmpRect.Top + 2, _
                            UIControl.Left + tmpRect.Left + UIControl.Width - 2, _
                            UIControl.Top + tmpRect.Top + UIControl.Height - 2, _
                            10, 10
            .setDrawStyle 0
        End If
        

        
        .DrawText UIControl.Left + tmpRect.Left, _
                        UIControl.Top + tmpRect.Top, _
                        Caption, False
    End With
End Sub
