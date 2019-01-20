VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Block"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type POINT
    ptX As Integer
    ptY As Integer
    isSingle As Boolean
    withPoints() As Variant
End Type

Private pt As POINT

 

Public Property Let PointY(ByVal vData As Integer)
    pt.ptY = vData
End Property

 
Public Property Get PointY() As Integer
    PointY = pt.ptY
End Property



Public Property Let PointX(ByVal vData As Integer)
    'MsgBox "Let Function Is Called "
 
    If Not (mvarMainPictureBox Is Nothing) Then
        Call DrawRect(mvarMainPictureBox, pt, vbWhite)
       ' intCoordinates(pt.ptX, pt.ptY) = 0
        pt.ptX = vData
        If vData > 0 Then
            If IsBound(pt) Then pt.ptX = pt.ptX - 1
        Else
            If IsBound(pt) Then pt.ptX = pt.ptX + 1
        End If
        Debug.Print "vData " & vData
        Debug.Print "Point " & pt.ptX
        
        intCoordinates(pt.ptX, pt.ptY) = 1
        Call DrawRect(mvarMainPictureBox, pt, vbRed)
    Else
        Error_001
    End If
End Property


Public Property Get PointX() As Integer
   ' MsgBox "Get Function Is Called "
   
    If IsBound(pt) Then
        PointX = pt.ptX - 1
    Else
        PointX = pt.ptX
    End If
    
    'Call DrawRect(mvarMainPictureBox, pt, vbRed)
End Property


Private Function IsBound(testedPoint As POINT) As Boolean
    
    If testedPoint.isSingle = True Then
        If _
        intCoordinates(testedPoint.ptX, testedPoint.ptY) <> 0 Or _
        intCoordinates(testedPoint.ptX, testedPoint.ptY) <> 0 Then _
        IsBound = True
        
    Else
        Dim pts() As Variant
        pts = testedPoint.withPoints
    End If
End Function

Private Sub DrawRect(whichObject As Object, pt As POINT, color As Long)
 
        hBrush = CreateSolidBrush(color)
        Range.startX = pt.ptX * deltaWidth + 1
        Range.startY = pt.ptY * deltaHeight + 1
        Range.endX = (pt.ptX + 1) * deltaWidth - 1
        Range.endY = (pt.ptY + 1) * deltaHeight - 1
        Call FillRect(whichObject.hdc, Range, hBrush)
        Call DeleteObject(hBrush)
End Sub


Private Sub Class_Initialize()
    pt.ptX = 3
    pt.ptY = 3
    pt.isSingle = True
    Call DrawRect(mvarMainPictureBox, pt, vbRed)
End Sub
