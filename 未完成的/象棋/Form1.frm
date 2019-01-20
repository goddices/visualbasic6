VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10005
   ClientLeft      =   1950
   ClientTop       =   825
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   667
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   735
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   9360
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   3000
   End
   Begin VB.ListBox List1 
      Height          =   420
      Left            =   9480
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   10005
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   665
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   598
      TabIndex        =   0
      Top             =   0
      Width           =   9000
      Begin VB.Image Img90 
         Height          =   1005
         Left            =   0
         Picture         =   "Form1.frx":2F3C
         Top             =   9000
         Width           =   1005
      End
      Begin VB.Image Image2 
         Height          =   1005
         Left            =   960
         Picture         =   "Form1.frx":38C8
         Top             =   6960
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image Image1 
         Height          =   1005
         Left            =   2640
         Picture         =   "Form1.frx":4254
         Top             =   7320
         Width           =   1005
      End
   End
   Begin VB.Image ImgEmpty 
      Height          =   1095
      Left            =   9720
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const deltaWIDTH = 67
Private Const DEGREEWIDTH = 9
Private Const DEGREEHEIGHT = 10

Private pic As Picture
Private pic2 As Picture

Private x_basic As Single
Private y_basic As Single
Private ptX As Integer
Private ptY As Integer
Private isMoveable As Boolean
Private Position(-1 To DEGREEWIDTH, -1 To DEGREEHEIGHT) As Integer

Private Sub Form_Activate()
    Dim I As Integer
    x_basic = Picture1.Width / DEGREEWIDTH
    y_basic = Picture1.Height / DEGREEHEIGHT
    Picture1.DrawWidth = 1
    Set pic = LoadPicture(App.Path + "\Ju.gif")
    Set pic2 = ImgEmpty.Picture
    Position(0, 9) = 10
    
    isMoveable = True
    
End Sub

Private Sub I_Click()

End Sub

Private Sub Image1_Click()
 MsgBox ptX & vbNewLine & ptY
 isMoveable = Not isMoveable
 MsgBox isMoveable
 If (Position(ptX, ptY) <> 0) And (Not isMoveable) Then
     Picture1.Cls
    Image2.Move ptX * deltaWIDTH, ptY * deltaWIDTH
    Timer1.Enabled = True
 End If
 
 If Position(ptX, ptY) = 0 And isMoveable Then
    Timer1.Enabled = False
    Position(ptX, ptY) = 10
    'Picture1.PaintPicture pic, ptX * deltaWIDTH, ptY * deltaWIDTH, deltaWIDTH, deltaWIDTH
    'Picture1.PaintPicture pic2, ptX * deltaWIDTH, ptY * deltaWIDTH, deltaWIDTH, deltaWIDTH
   '
 End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' a = Int(X / x_basic)
     b = Int(y / y_basic)

    MsgBox a & vbNewLine & b

    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        
        Dim a As Integer, b As Integer
        ptX = Int(x / x_basic)
        ptY = Int(y / y_basic)
        a = Int(x / x_basic) * deltaWIDTH
        b = Int(y / y_basic) * deltaWIDTH
        Image1.Move a, b
        List1.Clear
        List1.AddItem ptX & "   " & ptY
    
End Sub

Private Sub Dra()
   
End Sub



Private Sub Timer1_Timer()
Image2.Visible = Not Image2.Visible
End Sub
