VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUIManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const FORM_BASE_ID = 1024
Private Const FORM_CONTROL_MAX = 1024

Private Type RegFormInfo
    UIForm As CUIForm
    ID As Long
    '。。。
End Type

Private m_Forms() As RegFormInfo
Private m_FormCount As Long

Public Function LoadFormFormFile(FileName As String)
    Dim tmpForm As New CUIForm
    tmpForm.Create 200 * Rnd, 200 * Rnd, 320, 240
    RegWindow tmpForm
    Set tmpForm = Nothing
End Function

Public Function RegWindow(UIFrm As CUIForm) As Long
    RegWindow = FORM_BASE_ID + m_FormCount * FORM_CONTROL_MAX
    'New Form
    ReDim Preserve m_Forms(m_FormCount)
    With m_Forms(m_FormCount)
        Set .UIForm = UIFrm
        .ID = RegWindow
        .UIForm.ID = .ID
    End With
    m_FormCount = m_FormCount + 1
End Function

Public Sub SetFouce(UIFrm As CUIForm)
    m_Forms(m_FormCount - 1).UIForm.Actived = False
    
    Dim i As Integer, j As Integer, tmpFrmInfo As RegFormInfo
    
    For i = 0 To m_FormCount - 1
        If m_Forms(i).UIForm Is UIFrm Then
            '依次移动
            tmpFrmInfo = m_Forms(i)
            For j = i To m_FormCount - 2
                m_Forms(j) = m_Forms(j + 1)
            Next j
            Exit For
        End If
    Next i
    m_Forms(m_FormCount - 1) = tmpFrmInfo
    tmpFrmInfo.UIForm.Actived = True
End Sub

Public Sub Render()
    Dim i As Integer
    '排序渲染
    For i = 0 To m_FormCount - 1
        m_Forms(i).UIForm.Render
    Next i
End Sub

Public Sub SendInputMsg(MsgType As DefMsgType, ByVal Value As Long, ByVal x As Long, ByVal y As Long)
    Dim i As Integer
    If m_Forms(m_FormCount - 1).UIForm.Release Then
        '独占模式
        m_Forms(m_FormCount - 1).UIForm.SendInputMsg MsgType, Value, x, y
        Exit Sub
    End If
    Select Case MsgType
        Case MSG_MouseDown, MSG_MouseMove, MSG_MouseUp, MSG_MouseClick
            For i = m_FormCount - 1 To 0 Step -1
                If m_Forms(i).UIForm.IsInRect(x, y) Then
                    m_Forms(i).UIForm.SendInputMsg MsgType, Value, x, y
                    '切换窗体
                    If MsgType = MSG_MouseDown Then SetFouce m_Forms(i).UIForm
                    Exit For
                End If
            Next i
        Case MSG_KeyPress, MSG_Keydown, MSG_KeyUp
            '支持全局的热键
            m_Forms(m_FormCount - 1).UIForm.SendInputMsg MsgType, Value, x, y
    End Select
    
End Sub

Public Function GetForm(ByVal ID As Long) As CUIForm
    Dim i As Integer
    For i = 0 To m_FormCount - 1
        If ID = m_Forms(i).UIForm.ID Then
            Set GetForm = m_Forms(i).UIForm
            Exit Function
        End If
    Next
    Set GetForm = Nothing
End Function
