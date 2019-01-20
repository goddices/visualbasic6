VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBTetris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private objTargetPic1 As Object
Private objTargetPic2 As Object

Private objPoint As Point
Private objBar As Bar
Private objStartingBar As Bar
Private objNextBar As Bar

Private i As Integer, j As Integer
Private bStart As Boolean
Private intScore As Integer
Private intLoopTimes As Integer
 
  
Public Property Get Score() As Integer
    Score = intScore
End Property

 '// Byte   unsigned char
Public Property Get PointX() As Byte
     PointX = objPoint.ptX
End Property

Public Property Let PointX(getX As Byte)
    Call CreateBar(objTargetPic1, objPoint, objBar, True)
    If Not (IsBound(-1) Or IsBound(1)) Then objPoint.ptX = getX
    Call CreateBar(objTargetPic1, objPoint, objBar)
End Property

Public Property Get PointY() As Byte
    PointY = objPoint.ptY
End Property

Public Property Let PointY(getY As Byte)
    Call CreateBar(objTargetPic1, objPoint, objBar, True)
    objPoint.ptY = getY
    Call CreateBar(objTargetPic1, objPoint, objBar)
End Property

Public Property Get Transform() As Integer
    Transform = objBar.intTransform
End Property

Public Property Let Transform(getTransform As Integer)
    Call CreateBar(objTargetPic1, objPoint, objBar, True)
    objBar.intTransform = getTransform
    Call CreateBar(objTargetPic1, objPoint, objBar)
End Property

Public Property Get BarType() As String
    BarType = objBar.strBarType
End Property

Public Property Let BarType(getType As String)
    objBar.strBarType = getType
End Property

Private Function GetObject(getObj As Object) As Object
    GetObject = getObj
End Function
Public Property Let MainDrawingPictureBox(getObj As Object)
    
    If TypeOf getObj Is PictureBox Then
    
    'dlLenX = Int(getObj.Width / DLX)
    'dlLenY = Int(getObj.Height / DLY)
    Set objTargetPic1 = getObj
    
    objTargetPic1.BackColor = vbWhite
   ' Call CreateBar(objTargetPic1, objPoint, objBar)
    End If
End Property

 
 
Public Sub NextDrawingPictureBox(getObj As Object)
    objTargetPic2 = getObj
    
    Dim bpt As Point
    If bStart = True Then
        objBar = objStartingBar
        Call CreateBar(objTargetPic2, bpt, objNextBar, , True)
    
    End If
    Call CreateBar(objTargetPic2, bpt, objNextBar, , True)
End Sub

Private Sub StartGame()

    objPoint.ptX = 5
    objPoint.ptY = -1
    objStartingBar = RndType()
    objNextBar = RndType()
    bStart = True
     
    For i = 0 To 11
        For j = 0 To 15
            Coordinates(i, j) = 0
        Next
    Next
    
    For j = 0 To 16
        Coordinates(-1, j) = 1
        Coordinates(12, j) = 1
    Next
    
    For i = -1 To 12
        Coordinates(i, -1) = 0
        Coordinates(i, 16) = 1
    Next

End Sub

Private Sub BottomProgress()
    bStart = False
    Call CreateBar(objTargetPic1, objPoint, objBar)
    ''//LblScr.Caption = CStr(intScore)
    Call IsFullLine
    Call CreateNewBar
    intLoopTimes = 0
    Dim bpt As Point
    Pic2.Cls
    Call CreateBar(objTargetPic2, bpt, objBar, , True)

End Sub

Private Function IsTransformed(LPoint As Point, Optional BarType As String) As Boolean
    Dim intCorrect As Integer
    IsTransformed = True
    If BarType = "Line" Then
        intCorrect = 3
    Else
        intCorrect = 2
    End If
    
    For i = LPoint.ptX To LPoint.ptX + intCorrect
        For j = LPoint.ptY To LPoint.ptY + intCorrect
            If (Coordinates(i, j) <> 0) Then IsTransformed = False
        Next
    Next
End Function

Private Function IsBound(MoveStep As Integer) As Boolean
    IsBound = False
    For i = 0 To 7 Step 2
        If Coordinates(ptPointsContainer(i) + MoveStep, ptPointsContainer(i + 1)) <> 0 Then IsBound = True
    Next
End Function


Private Function IsBottom() As Boolean
    IsBottom = False
    
    For i = 0 To 7 Step 2
        If Coordinates(ptPointsContainer(i), ptPointsContainer(i + 1) + 1) <> 0 Then IsBottom = True
    Next

End Function

Private Sub IsFullLine()
'IsFullLine = False
'Dim IsFullLine_ As Boolean

    Dim bool  As Boolean
    Dim count As Integer
    bool = True
    For j = 15 To 0 Step -1
        For i = 0 To 11
            bool = bool And CBool(Coordinates(i, j))
        Next
        If bool = True Then
            count = count + 1
            Print j; count
            KillLines j, count
        End If
        bool = True
    Next

End Sub


Private Sub KillLines(LineNum As Integer, LinesCount As Integer)
    Dim pt_ As Point
    intLoopTimes = intLoopTimes + LinesCount
    
    intScore = intScore + 2 * intLoopTimes - 1
    '//LblScr.Caption = score
    
    pt_.ptX = 0
    pt_.ptY = LineNum
    'Print LineNum & " " & LinesCount
    Call KillingDraw(objTargetPic1, pt_, vbWhite)
    Call Translation(LineNum, LinesCount)

End Sub

Private Sub Translation(LineNum As Integer, MoveStep As Integer)
    Dim apt As Point
    
    For j = LineNum - MoveStep To 1 Step -1
        For i = 0 To 11
            If (Coordinates(i, j) <> 0 And j <> 0) Then
                Coordinates(i, j + MoveStep) = Coordinates(i, j)
                apt.ptX = i
                apt.ptY = j
                Call TranslatingDraw(objTargetPic1, apt, vbWhite)
                
                apt.ptY = j + MoveStep
                Call TranslatingDraw(objTargetPic1, apt, ColorB2L(Coordinates(i, j)))
                Coordinates(i, j) = 0
            End If
        Next
    Next
    Call IsFullLine
End Sub

Private Sub CreateNewBar()
    objBar = objNextBar
    objNextBar = RndType()
    objPoint.ptX = 5
    objPoint.ptY = -1
End Sub


Private Function RndType() As Bar
    Dim num As Integer
    Randomize
    num = Int(Rnd * 7)
    
    RndType.intTransform = 3
    Select Case num
        Case 0
            RndType.strBarType = "T"
            RndType.lBarColor = vbGreen
        Case 1
            RndType.strBarType = "L"
           RndType.lBarColor = vbBlue
        Case 2
            RndType.strBarType = "CL"
           RndType.lBarColor = vbMagenta
        Case 3
            RndType.strBarType = "Z"
           RndType.lBarColor = vbRed
        Case 4
            RndType.strBarType = "CZ"
           RndType.lBarColor = vbCyan
        Case 5
            RndType.strBarType = "B"
           RndType.lBarColor = vbGold
        Case 6
            RndType.strBarType = "Line"
           RndType.lBarColor = vbPurple1
    End Select
    
End Function



Private Sub Class_Initialize()
    objPoint.ptX = 5
    objPoint.ptY = -1
End Sub
