VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   11460
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
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

 
 
 
Private Sub Form_Click()
    Dim z As Complex
    Dim c As Complex
    Dim Red As Long, Grn As Long, Blu As Long
    Dim Escape As Boolean
    Dim x0 As Integer, y0 As Integer, i As Integer
    
    z.im = 0
    z.re = 0
    For x0 = -600 To 300
        For y0 = -200 To 200
            c.re = x0 / 200
            c.im = y0 / 200
            z.im = 0
            z.re = 0
            Escape = False
            For i = 0 To 255
                If Sqr(z.re * z.re + z.im * z.im) > 4 Then
                    Escape = True
                    Exit For
                Else
                    z = ComplexAdd(ComplexMuilt(z, z), c)
                End If
            Next
            
            If Escape Then
                Red = &HFF
                Grn = 0
                Blu = 0
            Else
                Red = i
                Grn = i
                Blu = i
            End If
            
            If Red > 255 Then
                Red = Red Mod 255
            End If
            
            If Grn > 255 Then
                Grn = Grn Mod 255
            End If
            
            If Blu > 255 Then
                Blu = Blu Mod 255
            End If
                
    
            SetPixel Me.hdc, x0 + 600, y0 + 200, RGB(Red, Grn, Blu)
            'Me.PSet (x0, y0), QBColor(i Mod 15)
        Next
        DoEvents
    Next
End Sub

 

 
