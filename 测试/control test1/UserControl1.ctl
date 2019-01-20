VERSION 5.00
Begin VB.UserControl TimerBar 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   ClipBehavior    =   0  '无
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   600
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   381
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "TimerBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


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
Private mColor As Long

Private ts As Boolean

Private Sub UserControl_Initialize()
    mRect.xs = 0
    mRect.ys = 0
    mRect.xe = Picture1.Width
    mRect.ye = Picture1.Height
    UserControl.Width = Picture1.Width
    UserControl.Height = Picture1.Height
End Sub

Private Sub FillPic(rec As RECT, color As Long)
    Dim hBrush As Long
    hBrush = CreateSolidBrush(color)
    FillRect Picture1.hdc, rec, hBrush
    DeleteObject hBrush
End Sub
 
 

Private Sub Timer1_Timer()
    Picture1.Cls
    mRect.xs = mRect.xs + 1
    If mRect.xe <= 0 Then Timer1.Enabled = False
    FillPic mRect, mColor
End Sub
 
Public Sub ReStart(Optional onoff As Boolean)
    mRect.xs = 0
    Timer1.Enabled = onoff
    ts = onoff
End Sub

Public Sub QuickStart(onoff As Boolean, color As Long, rate As Integer)
    Timer1.Enabled = onoff
    mColor = color
    If (rate <= 0 Or rate > 100) Then rate = 1
    Timer1.Interval = 100 / rate
End Sub

Public Property Let TimerSwitch(onoff As Boolean)
    ts = onoff
    If onoff = True Then
      Timer1.Enabled = True
    Else
      Timer1.Enabled = False
    End If
End Property

Public Property Get TimerSwitch() As Boolean
    TimerSwitch = ts
End Property


Public Property Let BarColor(color As Long)
    mColor = color
 
End Property
 
Public Property Get IsTimeUp() As Boolean
    If mRect.xs >= mRect.xe Then
        IsTimeUp = True
    Else
        IsTimeUp = False
    End If
End Property

Public Property Let TimerRate(rate As Integer)
    If (rate <= 0 Or rate > 100) Then rate = 1
    
    Timer1.Interval = 100 / rate
    
End Property

Public Property Get TimerRate() As Integer
    TimerRate = Timer1.Interval
End Property


