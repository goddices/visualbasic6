VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   7230
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1215
      Left            =   1320
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   3600
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const NUMLEN = 3
Private Sub Command1_Click()
Cls
Randomize
Dim i  As Integer, j As Integer
Dim num(NUMLEN - 1) As Integer
 
num(0) = Int(Rnd * NUMLEN + 1)

i = 1

10: Do

20:     num(i) = Int(Rnd * NUMLEN + 1)

25:     For j = 1 To i

30:         If num(i) = num(j - 1) Then GoTo 10

35:     Next

40:     i = i + 1

50: Loop Until i >= NUMLEN


For i = 0 To UBound(num)
   ' If num(i - 1) = num(i) Then MsgBox "two numbers are the same!": Exit For
    Print num(i)
Next
Print

End Sub

Private Sub Command2_Click()
Dim a(10) As Integer
MsgBox UBound(a)
For i = 0 To 10
a(i) = i
Next
End Sub
