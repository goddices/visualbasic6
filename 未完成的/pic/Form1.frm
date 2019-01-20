VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   503
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "to pic box"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   6840
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   360
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "load"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   6240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long _
) As Long

Private a As Picture

 
Private Sub Command1_Click()
        
        Set a = LoadPicture("test.bmp")
        
         Me.PaintPicture a, 40, 40, 44, 44, 0, 0, 44, 44
        'MsgBox Me.hDC
        
        
        
       ' Set x = Me.Picture
End Sub

Private Sub Command2_Click()
      BitBlt Picture1.hDC, 0, 0, 44, 44, Me.hDC, 40, 40, vbSrcCopy
End Sub
