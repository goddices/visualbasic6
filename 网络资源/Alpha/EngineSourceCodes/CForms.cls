VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mForms As New Collection
Private mActiveForm As CForm
Private mZOrder() As Integer

Public Sub PostMessages(MsgType As InputMsgType, ByVal Value As Long, ByVal x As Integer, ByVal y As Integer)
    '统一输入的控制
    Dim i As Integer
    If MsgType = MSG_MouseDown Then
        'First Treate The Mouse Pos to Check mActiveForm
        For i = mForms.Count To 1 Step -1
            If mForms.Item(i).CheckRect(x, y) Then ShowForm "", i: Exit For
        Next i
        If mActiveForm Is Nothing Then Exit Sub
        If Not mActiveForm.CheckRect(x, y) Then Set mActiveForm = Nothing
    End If
    If Not mActiveForm Is Nothing Then mActiveForm.GetMessages MsgType, Value
End Sub

Public Function Item(ByVal Name As String, Optional index As Integer) As CForm
    On Error GoTo errhand
    If index > 0 Then Set Item = mForms.Item(index) Else Set Item = mForms.Item(Name)
    Exit Function
errhand:
    Set Item = Nothing
End Function

Public Sub ShowForm(ByVal Name As String, Optional index As Integer)
    If Item(Name, index) Is Nothing Then Exit Sub
    If Not mActiveForm Is Nothing Then mActiveForm.Actived = False  'Lost Old HotPoint
    
    Set mActiveForm = Item(Name, index)
    mActiveForm.Visiable = True
    mActiveForm.Actived = True
    'exchange The Draw Order
    Dim i As Integer, s As Boolean
    For i = GetZorder(Name, index) To mForms.Count - 1
        mZOrder(i) = mZOrder(i + 1)
    Next i
    mZOrder(mForms.Count) = index
End Sub

Private Function GetZorder(ByVal Name As String, Optional index As Integer) As Integer
    Dim i As Integer
    '得到窗体的绘制次序
    If index > 0 Then
        For i = 1 To mForms.Count
            If mZOrder(i) = index Then GetZorder = i: Exit Function
        Next i
    Else
        For i = 1 To mForms.Count
            If mForms.Item(mZOrder(i)) Is Item(Name) Then GetZorder = i: Exit Function
        Next i
    End If
End Function


Public Function LoadForms() As Boolean
    '从文件中载入窗体
    Dim NewForm As New CForm
    Dim i As Integer
    On Error GoTo errhand
    For i = 1 To 5
        NewForm.Left = (i - 1) * 80
        NewForm.Right = (i - 1) * 80 + 100
        NewForm.Bottom = 100
        NewForm.Top = 10
        NewForm.Caption = Str(i)
        ReDim Preserve mZOrder(UBound(mZOrder) + 1)
        mZOrder(UBound(mZOrder)) = i
        mForms.Add NewForm
        Set NewForm = New CForm
    Next i
    Set mActiveForm = NewForm
    LoadForms = True
    Exit Function
errhand:
    
End Function

Private Sub Class_Initialize()
    ReDim mZOrder(0)
End Sub

Private Sub Class_Terminate()
    Set mForms = Nothing
End Sub

Public Sub Render()
    Dim i As Integer
    For i = 1 To mForms.Count
        mForms.Item(mZOrder(i)).Render
    Next i
End Sub
