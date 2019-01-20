VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   -150
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   221
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   381
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   2640
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type

Dim ptS  As POINTAPI
Dim ptE  As POINTAPI
Dim dwStartX As Long
Dim dwStartY As Long
Dim dwEndX As Long
Dim dwEndY As Long

Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOSIZE = &H1

Private Sub Command1_Click()
Me.Cls

Dim hwnd1 As Long
Dim hdc1 As Long
 hwnd1 = GetDesktopWindow()
 hdc1 = GetDC(hwnd1)
 Me.WindowState = vbMaximized
 SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE
 
 BitBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, hdc1, 0, 0, vbSrcCopy
 ReleaseDC hwnd1, hdc1
 
 
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

 

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
GetCursorPos ptS
dwStartX = ptS.x
dwStartY = ptS.y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
GetCursorPos ptE
dwEndX = ptE.x
dwEndY = ptE.y
End Sub

 
