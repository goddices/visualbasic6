VERSION 5.00
Begin VB.Form Fractal 
   Caption         =   "Fractal"
   ClientHeight    =   8280
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   13170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Zoom"
      Height          =   495
      Left            =   11040
      TabIndex        =   10
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Julia"
      Height          =   495
      Left            =   9120
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Zoom"
      Height          =   495
      Left            =   11040
      TabIndex        =   8
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mandelbrot"
      Height          =   495
      Left            =   9120
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   9120
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   9120
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtExp 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   9600
      TabIndex        =   3
      Text            =   "2"
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   11040
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picMan 
      AutoRedraw      =   -1  'True
      Height          =   7500
      Left            =   600
      ScaleHeight     =   500
      ScaleMode       =   0  'User
      ScaleWidth      =   500
      TabIndex        =   0
      Top             =   360
      Width           =   7500
   End
   Begin VB.Label Label2 
      Caption         =   "+C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Z=Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Fractal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim xMax As Single
Dim xMin As Single
Dim yMax As Single
Dim yMin As Single

Dim ZoomActive As Boolean
Dim Drawing As Boolean
Dim OldRect As Boolean

Dim x1 As Single
Dim x2 As Single
Dim y1 As Single
Dim y2 As Single

Private Sub StartPlot()

    Dim xmint As Single
    Dim xmaxt As Single
    Dim ymint As Single
    Dim ymaxt As Single

    xmint = xMin + (xMax - xMin) * x1 / picMan.ScaleWidth
    xmaxt = xMin + (xMax - xMin) * x2 / picMan.ScaleWidth
    
    ymint = yMin + (yMax - yMin) * y1 / picMan.ScaleHeight
    ymaxt = yMin + (yMax - yMin) * y2 / picMan.ScaleHeight
    
    xMax = xmaxt
    xMin = xmint
    yMax = ymaxt
    yMin = ymint
    
    Call PlotMan(xMin, xMax, yMin, yMax)
    
    OldRect = False
    
End Sub
Private Sub StartPlot1()

    Dim xmint As Single
    Dim xmaxt As Single
    Dim ymint As Single
    Dim ymaxt As Single

    xmint = xMin + (xMax - xMin) * x1 / picMan.ScaleWidth
    xmaxt = xMin + (xMax - xMin) * x2 / picMan.ScaleWidth
    
    ymint = yMin + (yMax - yMin) * y1 / picMan.ScaleHeight
    ymaxt = yMin + (yMax - yMin) * y2 / picMan.ScaleHeight
    
    xMax = xmaxt
    xMin = xmint
    yMax = ymaxt
    yMin = ymint
    
    Call PlotJulia(xMin, xMax, yMin, yMax)
    
    OldRect = False
    
End Sub

Private Sub RestoreStart()

    xMax = 2
    xMin = -2
    yMax = 2
    yMin = -2

    x1 = 0
    y1 = 0
    
    x2 = picMan.ScaleWidth
    y2 = x2 'picMan.ScaleHeight

End Sub
Private Sub PlotMan( _
                    ByVal xMin As Single, _
                    ByVal xMax As Single, _
                    ByVal yMin As Single, _
                    ByVal yMax As Single _
                    )

    Dim X As Single
    Dim Y As Single
    Dim k, c As Long
    
    Dim iMax As Integer
    Dim jMax As Integer
    Dim j As Integer
    Dim i As Integer
    
    Dim yDiff As Single
    Dim xDiff As Single
    
    Dim yConstant As Single
    Dim xConstant As Single
    Dim sglExp As Single
    Dim intRed, intGrn, intBlu   As Integer
    
    Drawing = True
    
    sglExp = txtExp.Text
    iMax = picMan.ScaleWidth
    jMax = iMax
    
    yDiff = yMax - yMin
    xDiff = xMax - xMin
    
    yConstant = yDiff / jMax
    xConstant = xDiff / iMax
    
    For j = 0 To jMax - 1
    
        Y = yMin + yConstant * j
        
        For i = 0 To iMax - 1
            X = xMin + xConstant * i
            k = Mandelbrot(X, Y, sglExp)
'            If k = 1000 Then
'                k = Mandelbrot(X, Y, -sglExp)
'            End If
            
            intRed = k * 6 Mod 256
            intGrn = k * 2 Mod 256
            intBlu = k * 13 Mod 256
            picMan.ForeColor = RGB(intRed, intGrn, 256 - intBlu)
            picMan.PSet (i, j)
        Next i
        
        picMan.Refresh
        DoEvents
        
    Next j

    Drawing = False
    
End Sub
Private Sub PlotJulia( _
                    ByVal xMin As Single, _
                    ByVal xMax As Single, _
                    ByVal yMin As Single, _
                    ByVal yMax As Single _
                    )

    Dim X As Single
    Dim Y As Single
    Dim k, c As Long
    
    Dim iMax As Integer
    Dim jMax As Integer
    Dim j As Integer
    Dim i As Integer
    
    Dim yDiff As Single
    Dim xDiff As Single
    
    Dim yConstant As Single
    Dim xConstant As Single
    Dim sglExp As Single
    Dim intRed, intGrn, intBlu   As Integer
    
    Drawing = True
    
    sglExp = txtExp.Text
    iMax = picMan.ScaleWidth
    jMax = iMax
    
    yDiff = yMax - yMin
    xDiff = xMax - xMin
    
    yConstant = yDiff / jMax
    xConstant = xDiff / iMax
    
    For j = 0 To jMax - 1
    
        Y = yMin + yConstant * j
        
        For i = 0 To iMax - 1
            X = xMin + xConstant * i
            k = Julia(X, Y, sglExp)
            intRed = k * 6 Mod 256
            intGrn = k * 6 Mod 256
            intBlu = k * 13 Mod 256
            picMan.ForeColor = RGB(256 - intRed, 256 - intGrn, 256 - intBlu)
            picMan.PSet (i, j)
        Next i
        
        picMan.Refresh
        DoEvents
        
    Next j

    Drawing = False
    
End Sub

Private Sub Command1_Click()
    Call RestoreStart
    
    picMan.Cls

    Call StartPlot
'    MsgBox ("ok")
End Sub

Private Sub Command2_Click()

Dim X, Y As Integer
Dim i, j, Red, Grn, Blu As Integer
Dim a, b, c, d, c2, d2, aa, bb, l, k As Double
Dim blnConverge As Boolean


    For c = -2# To 2# Step 0.001
        For d = -1.5 To 1.5 Step 0.001
            a = 0
            b = 0
            c2 = 0
            d2 = 0
            
            'Julia Set(, 0.188887129043845954792
            aa = c ^ 2 - d ^ 2
            bb = 2 * c * d
            For k = 1 To 1000
                a = aa + 0.285
                b = bb + 0.01
                If a ^ 2 + b ^ 2 > 4 Then
                    Exit For
                End If
                aa = a ^ 2 - b ^ 2
                bb = 2 * a * b
            
            Next
            Red = k * 9 Mod 256
            Grn = k * 1 Mod 256
            Blu = k * 9 Mod 256
            picMan.ForeColor = RGB(255 - Red, 255 - Grn, 255 - Blu)
            picMan.PSet (c * 5000 + 3750, d * 5000 + 3750)
        Next
    Next
    
    SavePicture picMan.Image, "C:\J4" & ".BMP"

End Sub

Private Sub Command3_Click()
Dim a, b, c, d, c2, d2, aa, bb, l, k, intExp, dblAng As Double
intExp = txtExp.Text

            aa = -1
            bb = -4
            dblAng = Atn(bb / aa)
            If aa > 0 And bb < 0 Then
                dblAng = 3.14159265358979 * 2 + Atn(bb / aa)
            End If
            
            If aa < 0 Then
                dblAng = 3.1415926535897 + Atn(bb / aa)
            End If
            c2 = (aa * aa + bb * bb) ^ (intExp / 2) * Cos(dblAng * intExp)
            d2 = (aa * aa + bb * bb) ^ (intExp / 2) * Sin(dblAng * intExp)
'            a = aa * aa - bb * bb
'            b = 2 * aa * bb
            
'         z = z ^ 3 + c
            a = aa ^ 3 - 3 * aa * bb ^ 2
            b = 3 * aa ^ 2 * bb - bb ^ 3
            
            MsgBox (CStr(c2) & "," & CStr(d2) & "," & CStr(a) & "," & CStr(b) & "," & CStr(dblAng))

End Sub

Private Sub Command4_Click()
    Dim Zr As Double
    Dim Zi As Double
    Dim dblZtr As Double
    Dim dblZti As Double
    Dim Zr2 As Double
    Dim Zi2 As Double
    Dim dblAng As Double
    
    Zr = 2
    Zi = 1
        
        Zr2 = (Zr * Zr - Zi * Zi) / ((Zr * Zr - Zi * Zi) ^ 2 + 4 * Zr * Zr * Zi * Zi)
        Zi2 = -2 * Zr * Zi / ((Zr * Zr - Zi * Zi) ^ 2 + 4 * Zr * Zr * Zi * Zi)
    MsgBox (CStr(Zr2) & "," & CStr(Zi2))

End Sub


Private Sub Command5_Click()
    If Drawing Then
        Beep
        Exit Sub
    End If

    picMan.Cls

    Call StartPlot

End Sub

Private Sub Command6_Click()
    Call RestoreStart
    
    picMan.Cls

    Call StartPlot1

End Sub

Private Sub Command7_Click()
    If Drawing Then
        Beep
        Exit Sub
    End If

    picMan.Cls

    Call StartPlot1

End Sub

Private Sub Form_Load()
'picMan.ScaleMode = 3

End Sub
Private Function Sqrt(ByVal Z As Double) As Double
    Sqrt = Z ^ 0.5
End Function

Private Function Mandelbrot(ByVal X As Double, ByVal Y As Double, ByVal EXP As Double) As Long

    Dim lngCount As Long
    Dim dblZr As Double
    Dim dblZi As Double
    Dim dblZtr As Double
    Dim dblZti As Double
    Dim dblZr2 As Double
    Dim dblZi2 As Double
    Dim dblAng As Double
    Dim dblExp As Double
    
    lngCount = 0
    dblZr = 0
    dblZi = 0
    dblExp = EXP
    
    Do Until lngCount >= 1024 Or (dblZr * dblZr + dblZi * dblZi) > 16
        dblZtr = dblZr2 + X
        dblZti = dblZi2 + Y
        dblZr = dblZtr
        dblZi = dblZti
        
        
        If dblZr = 0 Then
            If dblZi > 0 Then
                dblAng = 3.14159265358979 / 2
            Else
                dblAng = 3.14159265358979 * 1.5
            End If
        Else
            dblAng = Atn(dblZi / dblZr)
        End If
        If dblZr > 0 And dblZi < 0 Then
            dblAng = 3.14159265358979 * 2 + Atn(dblZi / dblZr)
        End If
        If dblZr < 0 Then
            dblAng = 3.1415926535897 + Atn(dblZi / dblZr)
        End If
        
'        dblExp = dblExp * (-1)
'        If dblExp < 0 Then
'            dblExp = dblExp + 2
'        End If
'        dblZr2 = (dblZr * dblZr - dblZi * dblZi) / ((dblZr * dblZr - dblZi * dblZi) ^ 2 + 4 * dblZr * dblZr * dblZi * dblZi)
'        dblZi2 = -2 * dblZr * dblZi / ((dblZr * dblZr - dblZi * dblZi) ^ 2 + 4 * dblZr * dblZr * dblZi * dblZi)
        dblZr2 = (dblZr * dblZr + dblZi * dblZi) ^ (dblExp / 2) * Cos(dblAng * dblExp)
        dblZi2 = (dblZr * dblZr + dblZi * dblZi) ^ (dblExp / 2) * Sin(dblAng * dblExp)
'        If dblExp < 0 Then
'            dblExp = dblExp - 2
'        End If
                
        lngCount = lngCount + 1
    Loop
    
    
    Mandelbrot = lngCount

End Function
Private Function Julia(ByVal X As Double, ByVal Y As Double, ByVal EXP As Double) As Long

    Dim lngCount As Long
    Dim dblZr As Double
    Dim dblZi As Double
    Dim dblZtr As Double
    Dim dblZti As Double
    Dim dblZr2 As Double
    Dim dblZi2 As Double
    Dim dblAng As Double
    
    lngCount = 0
    dblZr = 0
    dblZi = 0
        
        If X = 0 Then
            If Y > 0 Then
                dblAng = 3.14159265358979 / 2
            Else
                dblAng = 3.14159265358979 * 1.5
            End If
        Else
            dblAng = Atn(Y / X)
        End If
        If X > 0 And Y < 0 Then
            dblAng = 3.14159265358979 * 2 + Atn(Y / X)
        End If
        If X < 0 Then
            dblAng = 3.1415926535897 + Atn(Y / X)
        End If
        
        dblZr2 = (X * X + Y * Y) ^ (EXP / 2) * Cos(dblAng * EXP)
        dblZi2 = (X * X + Y * Y) ^ (EXP / 2) * Sin(dblAng * EXP)
    
    
    
    Do Until lngCount >= 1024 Or (dblZr * dblZr + dblZi * dblZi) > 4
        dblZtr = dblZr2 - 0.5
        dblZti = dblZi2 - 0.05
        dblZr = dblZtr
        dblZi = dblZti
        
        
        If dblZr = 0 Then
            If dblZi > 0 Then
                dblAng = 3.14159265358979 / 2
            Else
                dblAng = 3.14159265358979 * 1.5
            End If
        Else
            dblAng = Atn(dblZi / dblZr)
        End If
        If dblZr > 0 And dblZi < 0 Then
            dblAng = 3.14159265358979 * 2 + Atn(dblZi / dblZr)
        End If
        If dblZr < 0 Then
            dblAng = 3.1415926535897 + Atn(dblZi / dblZr)
        End If
        
        dblZr2 = (dblZr * dblZr + dblZi * dblZi) ^ (EXP / 2) * Cos(dblAng * EXP)
        dblZi2 = (dblZr * dblZr + dblZi * dblZi) ^ (EXP / 2) * Sin(dblAng * EXP)
                
        lngCount = lngCount + 1
    Loop
            
    Julia = lngCount

End Function


Private Sub picMan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Drawing Then
        Beep
        Exit Sub
    End If
        
    ZoomActive = True
    
    If OldRect Then
        picMan.DrawMode = vbInvert
        picMan.Line (x1, y1)-(x2, y2), RGB(0, 0, 0), B
        picMan.DrawMode = vbCopyPen
    End If
    
    x2 = X
    y2 = Y
    x1 = X
    y1 = Y
    
    picMan.DrawMode = vbInvert
    picMan.Line (x1, y1)-(x2, y2), RGB(0, 0, 0), B
    picMan.DrawMode = vbCopyPen
    
    OldRect = True
        
End Sub

Private Sub picMan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ZoomActive And Not Drawing Then Call AdjustBox(X, Y)
    
End Sub

Private Sub picMan_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ZoomActive And Not Drawing Then Call AdjustBox(X, Y)

    ZoomActive = False
    
End Sub

Private Sub AdjustBox(ByVal X As Single, ByVal Y As Single)

    Dim xDiff As Single
    Dim yDiff As Single
    
    xDiff = X - x1
    yDiff = Y - y1
    
    If xDiff <= 0 And yDiff <= 0 Then
        X = x1
        Y = y1
    End If
    
    picMan.DrawMode = vbInvert
    picMan.Line (x1, y1)-(x2, y2), RGB(0, 0, 0), B
    
    If xDiff > yDiff Then
        x2 = X
        y2 = y1 + (x2 - x1)
    ElseIf yDiff > xDiff Then
        y2 = Y
        x2 = x1 + (y2 - y1)
    End If

    picMan.Line (x1, y1)-(x2, y2), RGB(0, 0, 0), B
    picMan.DrawMode = vbCopyPen

End Sub


'Private Sub Command1_Click()
'Dim X, Y As Integer
'Dim i, j, Red, Grn, Blu  As Integer
'Dim a, b, c, d, e, c2, d2, aa, bb, l, k, intExp, dblAng As Double
'Dim blnConverge As Boolean
'intExp = txtExp.Text
'
''For e = 19 To 31
''intExp = e / 10#
'For c = -2 To 2# Step 0.005
'    For d = -1.5 To 1.5 Step 0.005
'        a = 0
'        b = 0
'            aa = 0
'            bb = 0
'            c2 = 0
'            d2 = 0
'        For k = 1 To 1000
'            a = c2 + c
'            b = d2 + d
'            If aa ^ 2 + bb ^ 2 > 4 Then
'                Exit For
'            End If
'            aa = a
'            bb = b
''         z = z ^ 2 + c
''            c2 = aa * aa - bb * bb
''            d2 = 2 * aa * bb
'            If aa = 0 Then
'                If bb > 0 Then
'                    dblAng = 3.14159265358979 / 2
'                Else
'                    dblAng = 3.14159265358979 * 1.5
'                End If
'            Else
'                dblAng = Atn(bb / aa)
'            End If
'
'            If aa > 0 And bb < 0 Then
'                dblAng = 3.14159265358979 * 2 + Atn(bb / aa)
'            End If
'
'            If aa < 0 Then
'                dblAng = 3.1415926535897 + Atn(bb / aa)
'            End If
'            c2 = (aa * aa + bb * bb) ^ (intExp / 2) * Cos(dblAng * intExp)
'            d2 = (aa * aa + bb * bb) ^ (intExp / 2) * Sin(dblAng * intExp)
'
''         z = z ^ 3 + c
''            c2 = aa ^ 3 - 3 * aa * bb ^ 2
''            d2 = 3 * aa ^ 2 * bb - bb ^ 3
''         z=z^4+c
''            c2 = (aa ^ 2 - bb ^ 2) ^ 2 - 4 * aa ^ 2 * bb ^ 2
''            d2 = 4 * aa * bb * (aa ^ 2 - bb ^ 2)
' '        z=z^5+c
''            c2 = aa ^ 5 - 10 * aa ^ 3 * bb ^ 2 + 5 * aa * bb ^ 4
''            d2 = 5 * aa ^ 4 * bb - 10 * aa ^ 2 * bb ^ 3 + bb ^ 5
' '        z=z^6+c
''            c2 = aa ^ 6 - 15 * aa ^ 4 * bb ^ 2 + 15 * aa ^ 2 * bb ^ 4 - bb ^ 6
''            d2 = 6 * aa ^ 5 * bb - 20 * aa ^ 3 * bb ^ 3 + 6 * aa * bb ^ 5
' '        z=z^7+c
''            c2 = aa ^ 7 - 21 * aa ^ 5 * bb ^ 2 + 35 * aa ^ 3 * bb ^ 4 - 7 * aa * bb ^ 6
''            d2 = 7 * aa ^ 6 * bb - 35 * aa ^ 4 * bb ^ 3 + 21 * aa ^ 2 * bb ^ 5 - bb ^ 7
'' '        z=z^0.5+c
''            If bb <> 0 Then
''            c2 = (aa * Sqrt(-aa / 2 + 1 / 2 * Sqrt(aa ^ 2 + bb ^ 2)) + Sqrt(aa ^ 2 + bb ^ 2) * Sqrt(-aa / 2 + 1 / 2 * Sqrt(aa ^ 2 + bb ^ 2))) / bb
''            d2 = Sqrt(-aa / 2 + 1 / 2 * Sqrt(aa ^ 2 + bb ^ 2))
''            Else
''
''            End If
'
'       Next
'
'
''    c2 = 0
''    d2 = 0
''        Zr = 0
''        Zi = 0
''    Do Until l >= 1000 Or (c2 + d2) > 40
''        Ztr = c2 - d2 + c
''        Zti = 2 * Zr * Zi + d
''        c2 = Zr * Zr
''        d2 = Zi * Zi
''        l = l + 1
''    Loop
''    Do Until Count >= 1000 Or (Zr2 + Zi2) > 4
''        Ztr = Zr2 - Zi2 + x
''        Zti = 2 * Zr * Zi + y
''        Zr = Ztr
''        Zi = Zti
''        Zr2 = Zr * Zr
''        Zi2 = Zi * Zi
''        Count = Count + 1
''    Loop
'
''            picMan.ForeColor = RGB(l / 1000# * 255, 0, 255)
''            picMan.PSet (c * 1250 + 2500, d * 1250 + 2500)
'            Red = k * 3 Mod 256
'            Grn = k * 2 Mod 256
'            Blu = k * 9 Mod 256
'            picMan.ForeColor = RGB(256 - Red, 256 - Grn, Blu)
'
'
''        If d > 0 Then
'            picMan.PSet (c * 2500 + 3725, d * 2500 + 3725)
''              picMan.PSet (c * 2500 + 3725, (2.25 - d ^ 2) ^ 0.5 * 2500), (1000 - k) * (1000 - k)
''        Else
'''            picMan.PSet (c * 2500 + 3725, (2.25 + d ^ 2) ^ 0.5 * 2500), (1000 - k) * (1000 - k)
''
''        End If
''
''
'''             picMan.PSet (c * 2500 + 3725, d * 2500 + 3725), (1000 - k) * (1000 - k)
''''要想得到彩色图形，最简单的方法是用迭代返回值n来着颜色。要想获得较好的艺术效果，一般对n做如下处理：
''''Red = n*Ar+Br;
''''Grn = n*Ag+Bg;
''''Blu = n*Ab+Bb;
''''if ((Red & 0x1FF) > 0xFF) Red = Red ^ 0xFF;
''''if ((Grn & 0x1FF) > 0xFF) Grn = Grn ^ 0xFF;
''''if ((Blu & 0x1FF) > 0xFF) Blu = Blu ^ 0xFF;
''''其中: Ar?Ag?Ab及Br?Bg?Bb为修正量
''''获得的Red、Grn、Blu为RGB三基色，着色效果为周期变化，具有较强的艺术感染力，而且等位线也蕴藏在周期变化的色彩之中。
'
'        If blnConverge Then
'
'       '     picMan.ForeColor = RGB(255, l - 1, 0)
'        Else
''            picMan.ForeColor = RGB(l - 1, 255, 0)
''            picMan.PSet (c * 2500 + 3725, d * 2500 + 3725)
'
'        End If
'    Next
'Next
'
''SavePicture picMan.Image, "C:\Z" & CStr(e) & ".BMP"
''Next
'
'End Sub
