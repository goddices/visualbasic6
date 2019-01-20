VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'保持属性值的局部变量
''mode  pixel
 
Public mBlock As New Block

Private mvarIsShowGrid As Boolean '局部复制

Public Property Let IsShowGrid(ByVal vData As Boolean)
    Dim i As Integer, j As Integer
    If vData Then
        If Not (mvarMainPictureBox Is Nothing) Then
            For i = 1 To intMaxX
                mvarMainPictureBox.Line (i * deltaWidth, 0)-(i * deltaWidth, intMaxY * deltaHeight)
            Next
            
            For j = 1 To intMaxY
                mvarMainPictureBox.Line (0, j * deltaHeight)-(intMaxX * deltaWidth, j * deltaHeight)
            Next
        Else
            Error_001
        End If
    Else
        mvarMainPictureBox.Cls
    End If
    mvarIsShowGrid = vData
End Property


Public Property Get IsShowGrid() As Boolean
 
    IsShowGrid = mvarIsShowGrid
End Property

Public Property Let SecondaryPictureBox(ByVal vData As Object)
 
    Set mvarSecondaryPictureBox = vData
   
End Property

Public Property Let MainPictureBox(ByVal vData As Object)
 
    If TypeOf vData Is PictureBox Then
        Set mvarMainPictureBox = vData
        With mvarMainPictureBox
            .BackColor = vbWhite
            .ScaleMode = 3
            .Appearance = 0
            .Width = intMaxX * deltaWidth
            .Height = intMaxY * deltaHeight
            .ScaleWidth = intMaxX * deltaWidth
            .ScaleHeight = intMaxY * deltaHeight
         
        End With
        
    Else
        Error_002
    End If
End Property

Private Sub Class_Initialize()
    For i = 0 To intMaxX - 1
        For j = 0 To intMaxY - 1
            intCoordinates(i, j) = 0
        Next
    Next
    
    For j = -1 To intMaxY
        intCoordinates(-1, j) = 1
        intCoordinates(intMaxX, j) = 1
    Next
    
    For i = -1 To intMaxX
        intCoordinates(i, -1) = 1
        intCoordinates(i, intMaxY) = 1
    Next
    
    
End Sub
