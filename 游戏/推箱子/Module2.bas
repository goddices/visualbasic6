Attribute VB_Name = "Module2"
Private ptObject As POINT
Private intH As Integer, intV As Integer

Public Sub MainPictureBox(ByVal vData As Object)
    If TypeOf vData Is PictureBox Then
        Set mvarMainPictureBox = vData
        With mvarMainPictureBox
            .BackColor = vbWhite
            .ScaleMode = 3
            .Appearance = 0
            .Width = intMaxX * deltaWidth + 2
            .Height = intMaxY * deltaHeight + 2
            '.ScaleWidth = intMaxX * deltaWidth
            '.ScaleHeight = intMaxY * deltaHeight
        End With
    End If
End Sub


Public Sub InitBound()
    Dim ii As Integer
    Dim jj As Integer
 
    
    For jj = -1 To intMaxY
        intCoordinates(-1, jj) = 1
        intCoordinates(intMaxX, jj) = 1
    Next
    For ii = -1 To intMaxX
        intCoordinates(ii, -1) = 1
        intCoordinates(ii, intMaxY) = 1
    Next
End Sub


Public Sub LocateWall(block1 As BLOCK, ByVal x1 As Integer, ByVal x2 As Integer, ByVal y1 As Integer, ByVal y2 As Integer)
    block1.x1 = x1
    block1.x2 = x2
    block1.y1 = y1
    block1.y2 = y2
End Sub


Public Sub CreateWall(blockWall As BLOCK)
    Dim temp As POINT
    For i = blockWall.x1 To blockWall.x2
        For j = blockWall.y1 To blockWall.y2
            temp.ptX = i
            temp.ptY = j
            intCoordinates(i, j) = 1
            Call DrawRect(temp, vbRed)
        Next
    Next
End Sub


Public Sub StartUpController(ptCtrl As POINT, ByVal x1 As Integer, ByVal y1 As Integer)
    ptCtrl.ptX = x1
    ptCtrl.ptY = y1
End Sub


Public Sub StartUpObjects(ptObj As POINT, ByVal x1 As Integer, ByVal y1 As Integer)
    ptObj.ptX = x1
    ptObj.ptY = y1
End Sub

Public Sub StartUpTargets(ptTar As POINT, ByVal x1 As Integer, ByVal y1 As Integer)
    ptTar.ptX = x1
    ptTar.ptY = y1
End Sub

Public Sub Objects(ptObj As POINT)
    intCoordinates(ptObj.ptX, ptObj.ptY) = 3
    Call DrawRect(ptObj, vbBlue)
End Sub


Public Sub Controller(pt As POINT)
    intCoordinates(pt.ptX, pt.ptY) = 2
    Call DrawRect(pt, vbGreen)
End Sub


Public Sub ShowTargets(pt As POINT)
    intCoordinates(pt.ptX, pt.ptY) = 4
    Call DrawRect(pt, vbCyan)
End Sub


Public Sub Xfunc(pt As POINT)
    'Debug.Print intCoordinates(pt.ptX, pt.ptY)
    intCoordinates(pt.ptX, pt.ptY) = 0
    Call DrawRect(pt, vbWhite)
End Sub

Public Function IsTarget(pt As POINT, intDirection As Integer) As Boolean
    Call GetHV(intDirection)
    If intCoordinates(pt.ptX + intH, pt.ptY + intV) = 4 Then
        IsTarget = True
    Else
        IsTarget = False
    End If
End Function


Public Function IsObject(pt1 As POINT) As Boolean
    If intCoordinates(pt1.ptX, pt1.ptY) = 3 Then IsObject = True
End Function


Public Function IsBound(pt1 As POINT) As Boolean
    If intCoordinates(pt1.ptX, pt1.ptY) = 1 Or _
        intCoordinates(pt1.ptX, pt1.ptY) = 2 Or _
        intCoordinates(pt1.ptX, pt1.ptY) = 3 _
    Then IsBound = True
End Function

Public Function GetObject() As POINT
    GetObject = ptObject
End Function

Public Function FindObject(ptCtrl As POINT, intDirection As Integer) As Boolean
    Dim temp As POINT
    Dim mm As Integer
    Call GetHV(intDirection)
    temp.ptX = ptCtrl.ptX + intH
    temp.ptY = ptCtrl.ptY + intV
    If intCoordinates(temp.ptX, temp.ptY) = 3 Then
        ptObject = temp
        FindObject = True
    Else
        FindObject = False
    End If
End Function

Public Sub MoveIt(ptCtrl As POINT, ptObj As POINT, ByVal intDirection As Integer)
    Call GetHV(intDirection)
    ptCtrl.ptX = ptCtrl.ptX + intH
    ptCtrl.ptY = ptCtrl.ptY + intV
    If IsObject(ptCtrl) Then
        ptObj.ptX = ptObj.ptX + intH
        ptObj.ptY = ptObj.ptY + intV
        If IsBound(ptObj) Then
            ptObj.ptX = ptObj.ptX - intH
            ptObj.ptY = ptObj.ptY - intV
            ptCtrl.ptX = ptCtrl.ptX - intH
            ptCtrl.ptY = ptCtrl.ptY - intV
        End If
    ElseIf IsBound(ptCtrl) Then
        ptCtrl.ptX = ptCtrl.ptX - intH
        ptCtrl.ptY = ptCtrl.ptY - intV
    End If
End Sub

'----------------------------------------------------------------
'inline  is better ...
Private Sub DrawRect(pt As POINT, color As Long)
    Dim hBrush As Long
    hBrush = CreateSolidBrush(color)
    Range.startX = pt.ptX * deltaWidth + 1
    Range.startY = pt.ptY * deltaHeight + 1
    Range.endX = (pt.ptX + 1) * deltaWidth - 1
    Range.endY = (pt.ptY + 1) * deltaHeight - 1
    Call FillRect(mvarMainPictureBox.hdc, Range, hBrush)
    Call DeleteObject(hBrush)
End Sub
'----------------------------------------------------------------



Private Function IsEqual(pt1 As POINT, pt2 As POINT) As Boolean
    If (pt1.ptX = pt2.ptX) And (pt1.ptY = pt1.ptY) Then IsEqual = True
End Function

Private Sub GetHV(ByVal intDirection As Integer)
    Select Case intDirection '必须为方向键
        Case 37
            intH = -1
            intV = 0
        Case 38
            intH = 0
            intV = -1
        Case 39
            intH = 1
            intV = 0
        Case 40
            intH = 0
            intV = 1
    End Select
End Sub

Public Function ShowPT(pt As POINT) As String
ShowPT = "Point: " & pt.ptX & " " & pt.ptY
End Function

