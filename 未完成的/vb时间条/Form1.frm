VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   394
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   3120
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   360
      ScaleHeight     =   315
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type RECT
xs As Long
ys As Long 'Left和Top为矩形区域左上角坐标
xe As Long
ye As Long 'Right和Bottom为矩形区域右下角坐标
End Type

Private Declare Function CreateSolidBrush Lib "gdi32" ( _
    ByVal crColor As Long _
) As Long


Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Private Declare Function FillRect Lib "user32" ( _
    ByVal hdc As Long, _
    lpRect As RECT, _
    ByVal hBrush As Long _
) As Long

Private mRect As RECT
Private sRect As RECT

Private Sub FillPic(rec As RECT, color As Long)
 
    hBrush = CreateSolidBrush(color)
    FillRect Picture1.hdc, rec, hBrush
    DeleteObject hBrush
End Sub

Private Sub Command1_Click()
    FillPic mRect, vbBlack
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    mRect.xs = 0
    mRect.ys = 0
    mRect.xe = Picture1.Width
    mRect.ye = Picture1.Height
    sRect = mRect
End Sub

Private Sub Timer1_Timer()
    Picture1.Cls
    mRect.xe = mRect.xe - 1
    If mRect.xe <= 0 Then MsgBox "end": Timer1.Enabled = False
    FillPic mRect, vbRed
End Sub
 
