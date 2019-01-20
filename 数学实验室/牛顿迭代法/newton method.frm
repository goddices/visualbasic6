VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   9885
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const e = 2.718281828459
Private Const pi = 3.1415926535898
Private Const Ei = 0.01

Private root1 As Complex
Private root2 As Complex
Private root3 As Complex

Private zzz1 As Complex
Private zzz2 As Complex
Private op As Integer

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Type Complex
    re As Double  '实部
    im As Double  ' 虚部
End Type
'复数赋值
Private Function ComplexAssign(re As Double, im As Double) As Complex
    ComplexAssign.re = re
    ComplexAssign.im = im
End Function
'加法
Private Function ComplexAdd(z1 As Complex, z2 As Complex) As Complex
    ComplexAdd.re = z1.re + z2.re
    ComplexAdd.im = z1.im + z2.im
End Function
'乘法
Private Function ComplexMulti(z1 As Complex, z2 As Complex) As Complex
    ComplexMulti.re = z1.re * z2.re - z1.im * z2.im
    ComplexMulti.im = z1.im * z2.re + z1.re * z2.im
End Function
'减法
Private Function ComplexMinus(z1 As Complex, z2 As Complex) As Complex
    ComplexMinus.re = z1.re - z2.re
    ComplexMinus.im = z1.im - z2.im
End Function
'除法
Private Function ComplexDivide(z1 As Complex, z2 As Complex) As Complex
    ComplexDivide.re = (z1.re * z2.re + z1.im * z2.im) / (z2.re ^ 2 + z2.im ^ 2)
    ComplexDivide.im = (z1.im * z2.re - z1.re * z2.im) / (z2.re ^ 2 + z2.im ^ 2)
End Function
'幂
Private Function ComplexExp(z As Complex, exp As Integer) As Complex
    Dim i As Integer
    Dim t As Complex
    t = z
    For i = 1 To exp - 1
        t = ComplexMulti(t, z)
    Next
    ComplexExp = t
End Function
'模
Private Function ComplexM(z As Complex) As Double
    ComplexM = Sqr(z.re ^ 2 + z.im ^ 2)
End Function
'欧拉公式
Private Function EulerFormula(fi As Double) As Complex
    EulerFormula.re = Cos(fi)
    EulerFormula.im = Sin(fi)
End Function
'两个复数的距离
Private Function ComplexDistance(z1 As Complex, z2 As Complex) As Double
    ComplexDistance = Sqr((z1.re - z2.re) ^ 2 + (z1.im - z2.im) ^ 2)
End Function

Private Sub PrintComplex(z As Complex)
    Dim operator As String
    If z.im < 0 Then
        operator = " - "
    Else
        operator = " + "
    End If
    Print z.re & operator & (-z.im) & " i"
End Sub

Private Function ComplexToString(z As Complex) As String
Dim operator As String
    If z.im < 0 Then
        operator = " - "
    Else
        operator = " + "
    End If
    ComplexToString = z.re & operator & (-z.im) & " i"
End Function

Private Sub Form_Click()
    Dim t1 As Complex, t2 As Complex, z As Complex, num1 As Complex, num3 As Complex
    Dim m_of_root1 As Double, m_of_root2 As Double, m_of_root3 As Double
    Dim x0 As Double, y0 As Double, i    As Integer
    Dim RR As Long, GG As Long, BB As Long
    Dim rd As Long, gr As Long, bl As Long
    
    num1 = ComplexAssign(1, 0)
    num3 = ComplexAssign(3, 0)
    
    rd = 5
    gr = 70
    bl = 180
 
   DoEvents
    For x0 = -1.5 To 1.5 Step 0.004
        For y0 = -1.5 To 1.5 Step 0.004
            z.re = x0
            z.im = y0
            If x0 = 0 And y0 = 0 Then GoTo continue
    
            
            For i = 1 To 30
              '  If ComplexM(z) > 1 Then Exit For
                t1 = ComplexMinus(ComplexExp(z, 3), num1)
                t2 = ComplexMulti(num3, ComplexExp(z, 2))
                
                z = ComplexMinus(z, ComplexDivide(t1, t2))
                If ComplexDistance(z, root1) < Ei Then
                    RR = 225 + i
                    Exit For
                End If
                
                If ComplexDistance(z, root2) < Ei Then
                    GG = 225 + i
                    Exit For
                End If
                
                If ComplexDistance(z, root3) < Ei Then
                    BB = 225 + i
                    Exit For
                End If
            Next
            RR = RR Mod 255
            GG = GG Mod 255
            BB = BB Mod 255
            SetPixel Me.hdc, 200 * (x0) + 300, 200 * (y0) + 300, RGB(RR, GG, BB)
continue: Next
    Next
End Sub


Private Sub Form_Load()
    root1 = ComplexAssign(1, 0)
    root2 = EulerFormula(2 * pi / 3)
    root3 = EulerFormula(4 * pi / 3)
End Sub
