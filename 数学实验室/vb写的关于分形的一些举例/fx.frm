VERSION 5.00
Begin VB.Form fx 
   BackColor       =   &H80000009&
   Caption         =   "分形举例"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12015
   LinkTopic       =   "Form2"
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   801
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      BackColor       =   &H80000009&
      Height          =   6795
      Left            =   60
      ScaleHeight     =   449
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   1
      Top             =   540
      Width           =   11835
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "运行"
      Height          =   495
      Left            =   1020
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "fx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'看了一些分形的资料后，自己随意作了一些试验，程序也没有仔细考虑
'如果你又什么建议，欢迎给我来信：qljqp@sohu.com
'---------program  by qlj----------------------------
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Type Complex
    re As Double  '实部
    im As Double  ' 虚部
End Type
Private Function ComplexAdd(a As Complex, b As Complex) As Complex
    Dim c As Complex
    c.re = a.re + b.re
    c.im = a.im + b.im
    ComplexAdd = c
End Function
Private Function ComplexMuilt(a As Complex, b As Complex) As Complex
    Dim c As Complex
    c.re = a.re * b.re - a.im * b.im
    c.im = a.im * b.re + a.re * b.im
    ComplexMuilt = c
End Function
Private Sub DrawMandelbrot(c As Complex)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim Red As Integer
    Dim Grn As Integer
    Dim Blu As Integer
    Dim Ar As Integer, Br As Integer
    Dim Ag As Integer, Bg As Integer
    Dim Ab As Integer, Bb As Integer
    
    Dim z As Complex
    
    Ar = 2
    Br = 50
    Ag = 3
    Bg = 50
    Ab = 4
    Bb = 100
    
    
    For i = -300 To 300
        For j = -200 To 200
            z.re = i / 200
            z.im = j / 200
            For k = 1 To 200
                If Sqr(z.re * z.re + z.im * z.im) > 4 Then
                    Exit For
                Else
                    z = ComplexAdd(ComplexMuilt(z, z), c)
                End If
                
                Red = k * Ar + Br
                Grn = k * Ag + Bg
                Blu = k * Ab + Bb
'                If ((Red & &H1FF) > &HFF) Then Red = Red ^ &HFF
'                If ((Grn & &H1FF) > &HFF) Then Grn = Grn ^ &HFF
'                If ((Blu & &H1FF) > &HFF) Then Blu = Blu ^ &HFF
                If Red > 255 Then
                    Red = Red Mod 255
                End If
                If Grn > 255 Then
                    Grn = Grn Mod 255
                End If
                If Blu > 255 Then
                    Blu = Blu Mod 255
                End If
                
                
                'pic.PSet (i + 300, j + 200), QBColor(k Mod 16)
                SetPixel pic.hdc, i + 300, j + 200, RGB(Red, Grn, Blu)
            Next k
        Next j
        DoEvents
    Next i
End Sub

Private Sub Command1_Click()
   ' Fractal_interpolating
    'sierpinski
    'aa
    'Exit Sub
    Dim c As Complex
    Dim a As Double
    Dim b As Double
    Dim i As Integer
    
    a = InputBox("请输入1-7中的一个数")
    c.re = -0.75
    c.im = 0
    Select Case a
    Case 1
    c.re = -0.75
    c.im = 0
    Case 2
    c.re = 0.45
    c.im = -0.1428
    Case 3
    c.re = 0.285
    c.im = 0.01
    Case 4
    c.re = 0.285
    c.im = 0
    Case 5
    c.re = -0.8
    c.im = 0.156
    Case 6
    c.re = -0.835
    c.im = -0.2321
    Case 7
    c.re = -0.70176
    c.im = -0.3842
    
    End Select
    
    DrawMandelbrot c
End Sub

Private Sub aa()
''''吸引盆（就是Julia集）
Const c = 5
pic.Scale (-c, c)-(c, -c)
Dim a, b As Double
Dim p, q As Double
Dim X, Y As Double
Dim i As Long
Dim u, v As Double

a = 3
b = 0.72

For p = -c To c Step c / 600
    For q = -c To c Step c / 600
        X = p
        Y = q
        For i = 1 To 150
            u = a * Y * Sin(X)
            v = X * X - b
            X = u
            Y = v
            If Sqr(X * X + Y * Y) > 50 Then
                Exit For
            ElseIf i > 120 Then
                pic.PSet (p, q)
                Exit For
            End If
        Next i
    Next q
Next p

End Sub

Private Sub sierpinski()

Dim s(100, 100), t(100, 100), X(12), Y(12)

Dim a(3), b(3), c(3), d(3), e(3), f(3)
Dim i, j, k, n As Integer
a(1) = 0.5

a(2) = 0.5

a(3) = 0.5

b(1) = 0

b(2) = 0

b(3) = 0

c(1) = 0

c(2) = 0

c(3) = 0

d(1) = 0.5

d(2) = 0.5

d(3) = 0.5

e(1) = 1

e(2) = 1

e(3) = 50

f(1) = 1

f(2) = 50

f(3) = 1

For i = 1 To 100

t(1, i) = 1

t(i, 1) = 1

t(100, i) = 1

t(i, 100) = 1

Next i

For n = 2 To 8

For i = 1 To 100

For j = 1 To 100

If (t(i, j) = 1) Then

For k = 1 To 3

s(a(k) * i + b(k) * j + e(k), c(k) * i + d(k) * j + f(k)) = 1

Next k

End If

Next j

Next i

For i = 1 To 100

For j = 1 To 100

t(i, j) = s(i, j)

s(i, j) = 0

If (t(i, j) = 1) Then

ScaleMode = 3

PSet (150 + i, 100 + j), RGB(0, 255, 0)

End If

Next j

Next i

Next n

End Sub

''2）分形插值曲线:

Private Sub Fractal_interpolating()

Dim X(3), f(3), d(3) As Double, a(10), e(10), c(10), ff(10)

Dim xx, yy, newx, newy As Double
Dim b, k As Double
Dim n, i As Integer

X(0) = 0

X(1) = 40

X(2) = 55

X(3) = 100

f(0) = 0

f(1) = 20

f(2) = 40

f(3) = 15

d(1) = 0.4

d(2) = -0.3

d(3) = 0.6

For n = 1 To 3

b = X(3) - X(0)

a(n) = (X(n) - X(n - 1)) / b

e(n) = (X(3) * X(n - 1) - X(0) * X(n)) / b

c(n) = (f(n) - f(n - 1) - d(n) * (f(3) - f(0))) / b

ff(n) = (X(3) * f(n - 1) - X(0) * f(n) - d(n) * (X(3) * f(0) - X(0) * f(3))) / b

Next

xx = 0

yy = 0

i = 0

Do While i < 5000

k = Int(3 * Rnd - 0.0001) + 1

newx = a(k) * xx + e(k)

newy = c(k) * xx + d(k) * yy + ff(k)

xx = newx

yy = newy

ScaleMode = 3

PSet (150 + xx * 3, 100 + yy * 2), RGB(0, 255, 0)

i = i + 1

Loop

End Sub

''3）分形树曲线：

Private Sub Fractal_tree()

X = 0

Y = 0

i = 0

Do

v = Rnd

Select Case v

Case 0 To 0.05

X = 0

Y = 0.5 * Y

Case 0.05 To 0.45

u = 0.42 * X - 0.42 * Y

Y = 0.2 + 0.42 * X + 0.42 * Y

X = u

Case 0.45 To 0.85

u = 0.42 * X + 0.42 * Y

Y = 0.2 - 0.42 * X + 0.42 * Y

X = u

Case Else

X = 0.1 * X

Y = 0.1 * Y + 0.2

End Select

ScaleMode = 3

PSet (200 + X * 400, 250 - Y * 400), RGB(0, 250, 0)

i = i + 1

Loop Until i > 8000

End Sub

''4)混沌图形一例：

Private Sub Chaos()

X = 0

Y = 0

xa = 200

ya = 100

xb = 100

yb = 200

xc = 400

yc = 400

Form1.ScaleMode = 3

i = 0

Do While i < 10000

k = Rnd()

If (k < 0.9) And (k > 0.6) Then

X = (X + xc) / 2

Y = (Y + yc) / 2

PSet (X, Y), RGB(0, 255, 0)

End If

If (k > 0.3) And (k < 0.6) Then

X = (X + xa) / 2

Y = (Y + ya) / 2

PSet (X, Y), RGB(0, 255, 0)

End If

If (k > 0.1) And (k < 0.3) Then

X = (X + xb) / 2

Y = (Y + yb) / 2

PSet (X, Y), RGB(0, 255, 0)

End If

i = i + 1

Loop

End Sub

