VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menace Map editor"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   105
      Max             =   12
      TabIndex        =   5
      Top             =   2325
      Width           =   7035
   End
   Begin VB.PictureBox Picture2 
      Height          =   1800
      Left            =   120
      ScaleHeight     =   116
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   464
      TabIndex        =   4
      Top             =   450
      Width           =   7020
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   330
      Left            =   5175
      Picture         =   "map.frx":0000
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   126
      TabIndex        =   2
      Top             =   45
      Width           =   1950
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   270
      Left            =   1440
      Max             =   40
      Min             =   1
      TabIndex        =   1
      Top             =   75
      Value           =   1
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "Left             Right"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   105
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Current Map 1"
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mapnum%
Dim cell%(40, 8)
Dim lb%, rb%, dwn%

Private Sub Form_Load()
mapnum% = 1
loadmap
End Sub

Private Sub Form_Unload(Cancel As Integer)
savemap
End Sub

Private Sub HScroll1_Change()
Label3.Caption = "Current Map " & HScroll1.Value
If mapnum% > 0 Then savemap
mapnum% = HScroll1.Value
loadmap
End Sub

Private Sub HScroll2_Change()
Picture2_Paint
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        lb% = x \ 18
        u& = BitBlt(hdc, 243, 5, 18, 18, Picture1.hdc, lb% * 18, 0, SRCCOPY)
    ElseIf Button = 2 Then
        rb% = x \ 18
        u& = BitBlt(hdc, 292, 5, 18, 18, Picture1.hdc, rb% * 18, 0, SRCCOPY)
    End If
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
dwn% = True
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If dwn% Then
    x = x \ 18: y = y \ 18
    If Button = 1 Then
        cell%(x + HScroll2.Value, y) = lb%
        u& = BitBlt(Picture2.hdc, x * 18, y * 18, 18, 18, Picture1.hdc, cell%(x + HScroll2.Value, y) * 18, 0, SRCCOPY)
    Else
        cell%(x + HScroll2.Value, y) = rb%
        u& = BitBlt(Picture2.hdc, x * 18, y * 18, 18, 18, Picture1.hdc, cell%(x + HScroll2.Value, y) * 18, 0, SRCCOPY)
    End If
End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
dwn% = False
End Sub

Private Sub Picture2_Paint()
'29 wide '21 high
For a% = 0 To 28
    For b% = 0 To 5
        u& = BitBlt(Picture2.hdc, a% * 18, b% * 18, 18, 18, Picture1.hdc, cell%(a% + HScroll2.Value, b%) * 18, 0, SRCCOPY)
    Next
Next
u& = BitBlt(hdc, 243, 5, 18, 18, Picture1.hdc, lb% * 18, 0, SRCCOPY)
u& = BitBlt(hdc, 292, 5, 18, 18, Picture1.hdc, rb% * 18, 0, SRCCOPY)
End Sub
Sub loadmap()
Open App.Path & "\map.dat" For Random As #1 Len = 2
nfile% = LOF(1) / 2
For a% = 0 To 39
    For b% = 0 To 5
        Get #1, 1 + (a% + (b% * 40)) + (mapnum% - 1) * 2500, cell(a%, b%)
    Next
Next

Picture2_Paint
Close
End Sub
Sub savemap()
Open App.Path & "\map.dat" For Random As #1 Len = 2
For a% = 0 To 39
    For b% = 0 To 5
        Put #1, 1 + (a% + (b% * 40)) + (mapnum% - 1) * 2500, cell(a%, b%)
    Next
Next
Close
End Sub

Private Sub VScroll1_Change()
Picture2_Paint

End Sub
