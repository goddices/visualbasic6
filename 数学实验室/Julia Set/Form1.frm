VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   594
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   780
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   8280
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   390
      Left            =   1800
      TabIndex        =   1
      Text            =   "0"
      Top             =   8280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   "0"
      Top             =   8280
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Command1_Click()
    Dim c As Complex
    If IsNumeric(Text1.Text) And IsNumeric(Text2.Text) Then
        c.re = CDbl(Text1.Text)
        c.im = CDbl(Text2.Text)
        JuliaSet c
    End If
End Sub
 
 
Private Sub JuliaSet(z0 As Complex)
    Dim z As Complex
    Dim c As Complex
    
    Dim Ar As Integer, Br As Integer, Ag As Integer, Bg As Integer, Ab As Integer, Bb As Integer
    Ar = 2
    Br = 50
    Ag = 3
    Bg = 50
    Ab = 4
    Bb = 100
    
    
    c.re = z0.re
    c.im = z0.im
    For x0 = -400 To 400
        For y0 = -300 To 300
            z.re = x0 / 200
            z.im = y0 / 200
            For i = 1 To 200
                If Sqr(z.re * z.re + z.im * z.im) > 4 Then
                    Exit For
                Else
                    z = ComplexAdd(ComplexMuilt(z, z), c)
                End If
            Next
   
            Red = i * Ar + Br
            Grn = i * Ag + Bg
            Blu = i * Ab + Bb

            If Red > 255 Then
                Red = Red Mod 255
            End If
            
            If Grn > 255 Then
                Grn = Grn Mod 255
            End If
            
            If Blu > 255 Then
                Blu = Blu Mod 255
            End If
                
    
            SetPixel Me.hdc, x0 + 350, y0 + 250, RGB(Red, Grn, Blu)
            'Me.PSet (x0, y0), QBColor(i Mod 15)
        Next
        DoEvents
    Next
End Sub

 
