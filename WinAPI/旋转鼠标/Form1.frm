VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Haha"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   149
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   244
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2160
      Top             =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "关闭 Alt F4"
      Height          =   180
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Const PI = 3.1415926
Private Const r = 300
Private Const a = 500
Private Const b = 400
Private t As Integer, rad As Double
Private x As Long, y As Long
'
 
Private Sub Form_Activate()

SetCursorPos 2 * a, b
End Sub

Private Sub Timer1_Timer()
t = t + 1
If t = 360 Then t = 0
rad = PI / 180 * t
x = CLng(r * Cos(rad)) + a
y = CLng(r * Sin(rad)) + b
SetCursorPos x, y
End Sub
