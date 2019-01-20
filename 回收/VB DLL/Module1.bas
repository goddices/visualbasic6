Attribute VB_Name = "Module1"
Public Const DLX   As Integer = 12
Public Const DLY As Integer = 16

Public Const vbGold As Long = 55295              'RGB(255, 215, 0)
Public Const vbPurple1 As Long = 16724123        'RGB(155, 48, 255)

Public hBrush As Long

Public dlLenX As Integer, dlLenY As Integer  'Δx'Δy

Public Coordinates(-1 To DLX + 1, -1 To DLY + 1) As Byte

Public ptPointsContainer() As Variant

Public Type Point
    ptX As Integer
    ptY As Integer
End Type


Public Type Bar
    strBarType As String
    lBarColor As Long
    intTransform As Integer
End Type


Public Type RECT
    xs As Long
    ys As Long
    xe As Long
    ye As Long
End Type


Public Declare Function CreateSolidBrush Lib "gdi32" ( _
    ByVal crColor As Long _
) As Long


Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function FillRect Lib "user32" ( _
    ByVal hdc As Long, _
    lpRect As RECT, _
    ByVal hBrush As Long _
) As Long


Public Function ColorL2B(lColor As Long) As Byte
    Select Case lColor
        Case vbWhite
            ColorL2B = 0
        Case vbRed
            ColorL2B = 1
        Case vbCyan
            ColorL2B = 2
        Case vbGreen
            ColorL2B = 3
        Case vbBlue
            ColorL2B = 4
        Case vbMagenta
            ColorL2B = 5
        Case vbGold
            ColorL2B = 6
        Case vbPurple1
            ColorL2B = 7
    End Select
End Function

Public Function ColorB2L(bColor As Byte) As Long
    Select Case bColor
        Case 0
            ColorB2L = vbWhite
        Case 1
            ColorB2L = vbRed
        Case 2
            ColorB2L = vbCyan
        Case 3
            ColorB2L = vbGreen
        Case 4
            ColorB2L = vbBlue
        Case 5
            ColorB2L = vbMagenta
        Case 6
            ColorB2L = vbGold
        Case 7
            ColorB2L = vbPurple1
    End Select
End Function

Public Sub CreateBar(Object As PictureBox, spt As Point, thisBar As Bar, Optional isDown As Boolean, Optional isNext As Boolean)        'As Point
    
    Dim pt(1 To 4) As Point
    Dim X0 As Integer, Y0 As Integer, X1 As Integer, Y1 As Integer
    Dim X2 As Integer, Y2 As Integer, X3 As Integer, Y3 As Integer
    '//枚举表达式,为了方便
    X0 = spt.ptX + 0
    Y0 = spt.ptY + 0
    X1 = spt.ptX + 1
    Y1 = spt.ptY + 1
    X2 = spt.ptX + 2
    Y2 = spt.ptY + 2
    X3 = spt.ptX + 3
    Y3 = spt.ptY + 3
    
    
    '(X0,Y0)  (X1,Y0)  (X2,Y0)  (X3,Y0)
    '
    '(X0,Y1)  (X1,Y1)  (X2,Y1)  (X3,Y1)
    '
    '(X0,Y2)  (X1,Y2)  (X2,Y2)  (X3,Y2)
    '
    '(X0,Y3)  (X1,Y3)  (X2,Y3)  (X3,Y3)
    '
    '(X0,Y4)  (X1,Y4)  (X2,Y4)  (X3,Y4)
    '
    '
    '特殊方块，长直线   //////////              □
    '---------------                           '□
    X4 = spt.ptX + 4   '////////////□□□□    □
    Y4 = spt.ptY + 4                          ' □
    '---------------
    Dim bColorCode As Byte
    
    bColorCode = ColorL2B(thisBar.lBarColor)
    
    If isDown Then bColorCode = 0
    
    Select Case thisBar.strBarType & CStr(thisBar.intTransform)
    
           
        Case "T1"                        '□□□
                                         '  □
            ptPointsContainer = Array(X0, Y1, X1, Y1, X2, Y1, X1, Y2)
            
     
            
        Case "T2"                       '□
                                        '□□
                                        '□
            ptPointsContainer = Array(X1, Y0, X1, Y1, X1, Y2, X2, Y1)
            
     
            
        Case "T3"                       '  □
     
            
            ptPointsContainer = Array(X1, Y0, X0, Y1, X1, Y1, X2, Y1)
            
     
            
        Case "T4"                       '  □
    
            ptPointsContainer = Array(X0, Y1, X1, Y0, X1, Y1, X1, Y2)
            
     
        Case "L1"                       '    □
    
            ptPointsContainer = Array(X0, Y1, X1, Y1, X2, Y1, X2, Y0)
            
     
            
        Case "L2"                       ' □
    
    
            ptPointsContainer = Array(X1, Y0, X1, Y1, X1, Y2, X2, Y2)
            
     
            
        Case "L3"
    
    
            ptPointsContainer = Array(X0, Y1, X1, Y1, X2, Y1, X0, Y2)
            
     
            
        Case "L4"                       ' □□
    
    
            ptPointsContainer = Array(X0, Y0, X1, Y0, X1, Y1, X1, Y2)
            
     
            
        Case "CL1"                     ' □
    
            
            ptPointsContainer = Array(X0, Y0, X0, Y1, X1, Y1, X2, Y1)
            
     
    
        Case "CL2"                     '   □
    
    
            ptPointsContainer = Array(X1, Y0, X1, Y1, X1, Y2, X0, Y2)
            
     
            
            
        Case "CL3"                     '□□□
    
            
            ptPointsContainer = Array(X0, Y1, X1, Y1, X2, Y1, X2, Y2)
            
     
            
        
        Case "CL4"                    '□□
    
            
            ptPointsContainer = Array(X1, Y0, X2, Y0, X1, Y1, X1, Y2)
            
     
            
        Case "Z1", "Z3"                 '□□
    
            
            ptPointsContainer = Array(X0, Y0, X1, Y0, X1, Y1, X2, Y1)
            
     
            
        Case "Z2", "Z4"                  '  □
    
            
            ptPointsContainer = Array(X1, Y0, X1, Y1, X0, Y1, X0, Y2)
            
     
            
        Case "CZ1", "CZ3"               '  □□
    
           
            ptPointsContainer = Array(X1, Y0, X2, Y0, X0, Y1, X1, Y1)
     
            
        Case "CZ2", "CZ4"               '□
                                        '□□
    
        
            ptPointsContainer = Array(X0, Y0, X0, Y1, X1, Y1, X1, Y2)
            
     
            
        Case "B1", "B2", "B3", "B4"     '□□
    
            
            ptPointsContainer = Array(X0, Y0, X1, Y0, X0, Y1, X1, Y1)
    
     
            
        Case "Line1", "Line3"       '□
                                    '□
    
            
            ptPointsContainer = Array(X1, Y0, X1, Y1, X1, Y2, X1, Y3)
            
     
            
        Case "Line2", "Line4"
    
            
            ptPointsContainer = Array(X0, Y1, X1, Y1, X2, Y1, X3, Y1)
            
     
    End Select
            
    If Not isNext Then
        For i = 0 To 7 Step 2
            Coordinates(ptPointsContainer(i), ptPointsContainer(i + 1)) = ColorCode
        Next
    End If
    
    'Call DrawBar(Object, ptPointsContainer(), color)

End Sub

Public Sub DrawBar(Object As PictureBox, pdot() As Variant, color As Long)
    Dim Range As RECT
    hBrush = CreateSolidBrush(color)
    
    For i = 0 To UBound(pdot) Step 2
        Range.xs = pdot(i) * dlLenX + 1
        Range.ys = pdot(i + 1) * dlLenY + 1
        Range.xe = (pdot(i) + 1) * dlLenX - 1
        Range.ye = (pdot(i + 1) + 1) * dlLenY - 1
        FillRect Object.hdc, Range, hBrush
    Next
    DeleteObject (hBrush)
End Sub


Public Sub KillingDraw(Object As PictureBox, pt1 As Point, color As Long)
    Dim Range As RECT
    Dim ColorCode As Byte
    hBrush = CreateSolidBrush(color)
  

    ColorCode = ColorL2B(color)
    
    For i = 0 To DLX - 1
        Coordinates(i, pt1.ptY) = ColorCode
    Next
    
    For i = pt1.ptX To DLX - 1
        'Form1.Print i & "  "
        Range.xs = i * dlLenX + 1
        Range.ys = pt1.ptY * dlLenY + 1
        Range.xe = (i + 1) * dlLenX - 1
        Range.ye = (pt1.ptY + 1) * dlLenY - 1
        FillRect Object.hdc, Range, hBrush
    Next
    DeleteObject (hBrush)
End Sub

Public Sub TranslatingDraw(Object As PictureBox, pt1 As Point, color As Long)
    Dim Range As RECT
    hBrush = CreateSolidBrush(color)
    Range.xs = pt1.ptX * dlLenX + 1
    Range.ys = pt1.ptY * dlLenY + 1
    Range.xe = (pt1.ptX + 1) * dlLenX - 1
    Range.ye = (pt1.ptY + 1) * dlLenY - 1
    FillRect Object.hdc, Range, hBrush
    DeleteObject (hBrush)
End Sub
