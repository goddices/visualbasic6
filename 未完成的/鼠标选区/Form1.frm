VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Height          =   6015
      Left            =   480
      ScaleHeight     =   397
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   389
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Const GAME_PUL = 10
'Private Const GAME_WIDTH = 380
'Private Const GAME_HEIGHT = 350
Private startX As Single, startY As Single
Private endX As Single, endY As Single
'Private selectRegion  As Boolean
Private rdrw() As Variant


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Dim unitX As Single, unitY As Single
    
    If Button = 1 Then
        startX = X
        startY = Y
       ' unitStartX = Int(X / GAME_PUL)
        'unitStartY = Int(Y / GAME_PUL)
    
    ElseIf Button = 2 Then
    
        Static count As Integer
        count = count + 1
        Call FillArray(count, startX, startY, endX, endY)
        'MsgBox X & Y & vbNewLine & endX & endY
        
     '   If selectRegion = False Then
      '      selectRegion = True
            
       '     strPoint = strPoint & unitStartX & "," & unitStartY & "," & unitEndX & "," & unitEndY & vbNewLine
       '     selectRegion = False
      '  End If
    End If
 
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Picture1.Cls
        Call ReDraw
        Picture1.Line (startX, startY)-(X, Y), vbRed, B
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        endX = X
        endY = Y
        'unitEndX = Int(X / GAME_PUL)
        'unitEndY = Int(Y / GAME_PUL)
    End If
End Sub

Private Sub FillArray(ByVal i As Integer, ByVal sx As Integer, ByVal sy As Integer, ByVal ex As Integer, ByVal ey As Integer)
    
    ReDim Preserve rdrw(i)
    Print i
    Dim arr() As String
    rdrw(i) = CStr(sx) & "," & CStr(sy) & "," & CStr(ex) & "," & CStr(ey)
  
End Sub

Private Sub ReDraw()
    On Error Resume Next ''第一次使用下标越界
    For i = 1 To UBound(rdrw)
        arr = Split(rdrw(i), ",")
        Picture1.Line (arr(0), arr(1))-(arr(2), arr(3)), vbRed, B
    Next
End Sub

