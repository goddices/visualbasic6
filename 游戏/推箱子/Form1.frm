VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   5880
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private c1 As POINT ' Controller
'Private o2  As POINT ' Object 2
Private ot As POINT ' Object Tempary'
Private tar As POINT
Private bTarget(1) As Boolean
Private bGetToTarget As Boolean
Private Sub Command1_Click()
    Dim w1 As BLOCK
    Call LocateWall(w1, 0, 0, 0, 15)
    Call CreateWall(w1)
    Call StartUpController(c1, 7, 4)
    Call Controller(c1)
    Call StartUpObjects(o(0), 5, 7)
    Call Objects(o(0))
    Call StartUpObjects(o(1), 6, 8)
    Call Objects(o(1))
    Call StartUpTargets(tar, 5, 3)
    Call ShowTargets(tar)
    ot = o(1)
End Sub


 
'------------------------codes  ------------------------------------
     'Select Case KeyCode
       ' Case 37
          '  Call MoveIt(c1, o1, 37)
            'c1.ptX = c1.ptX - 1
            'If IsObject(c1) Then
            '    o1.ptX = o1.ptX - 1
            '    If IsBound(o1) Then
            '        o1.ptX = o1.ptX + 1
            '        c1.ptX = c1.ptX + 1
            '    End If
            'End If
    
       ' Case 38
           ' Call MoveIt(c1, o1, 38)
            'c1.ptY = c1.ptY - 1
            'If IsObject(c1) Then c1.ptY = c1.ptY + 1
            'If IsBound(c1) Then c1.ptY = c1.ptY + 1
        'Case 39
            'c1.ptX = c1.ptX + 1
            'If IsObject(c1) Then o1.ptX = o1.ptX - 1
            'If IsBound(c1) Then c1.ptX = c1.ptX - 1
       ' Case 40
            'c1.ptY = c1.ptY + 1
            'If IsBound(c1) Then c1.ptY = c1.ptY - 1
    'End Select
    'Debug.Print c1.ptX & "  " & c1.ptY
'-------------------------------------------------------------------



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Xfunc(c1)
    bTarget(0) = IsTarget(c1, KeyCode)
    bTarget(1) = IsTarget(ot, KeyCode)
    If FindObject(c1, KeyCode) Then ot = GetObject()
    Call MoveIt(c1, ot, KeyCode)
   ' Debug.Print bTarget(0)
   ' Debug.Print bTarget(1)
    If Not (bTarget(0) And bTarget(1)) Then Call ShowTargets(tar)
  '  Debug.Print "c1: " & ShowPT(c1)
  '  Debug.Print "o1: " & ShowPT(o(0))
  ''  Debug.Print "o2: " & ShowPT(o(1))
   ' Debug.Print "ot: " & ShowPT(ot)
    'Debug.Print bTarget(0) Or bTarget(1)
    'Debug.Print "ot " & ot.ptX & "  " & ot.ptY
   ' Debug.Print "targ " & tar.ptX & "  " & tar.ptY
   ' If bIsFind Or Not (bTarget(0) And bTarget(1)) Then
    Call Objects(ot)
    Call Controller(c1)
End Sub

Private Sub Form_Load()
    Call MainPictureBox(Picture1)
    Call InitBound '(void) why doesn't it exist ?
  
End Sub

 

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim str As String
    str = "( " + CStr(Int(X / 20)) + " , " + CStr(Int(Y / 20)) + " ) = "
    str = str + CStr(intCoordinates(Int(X / 20), Int(Y / 20)))
    Debug.Print str
End Sub




