VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   -150
   ClientTop       =   435
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   462
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   4800
      Left            =   5520
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   1500
      Width           =   4500
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   4800
      Left            =   120
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   1500
      Width           =   4500
      Begin VB.Image Image1 
         Height          =   720
         Left            =   2280
         MouseIcon       =   "Form1.frx":53E2
         Picture         =   "Form1.frx":A7C4
         Top             =   1800
         Width           =   720
      End
   End
   Begin VB.Label Label1 
      Caption         =   "还剩 不同"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCursorPos Lib "user32" ( _
    ByVal crX As Long, _
    ByVal crY As Long _
) As Long

Private Const GAME_WIDTH = 380
Private Const GAME_HEIGHT = 350
Private Const GAME_STARTLEFT = 10
Private Const GAME_TOP = 100
Private Const GAME_DIFF = 10
Private Const GAME_PUL = 10 ''PUL Per Unit Length
''

Private file1 As String
Private file2 As String
Private picCount As Integer
Private gameString() As String
Private ptpos() As Integer
Private diffCount As Integer
Private isOverSelect() As Boolean
Private rndNum()  As Integer

Private Sub DiffRndNum(ByVal LON As Integer) ' length of numbers

    Randomize
    
    Dim i  As Integer, j As Integer
    
    ReDim rndNum(LON - 1) As Integer
     
    rndNum(0) = Int(Rnd * LON + 1)
    
    i = 1

10: Do

20:     rndNum(i) = Int(Rnd * LON + 1)

25:     For j = 1 To i

30:         If rndNum(i) = rndNum(j - 1) Then GoTo 10

35:     Next

40:     i = i + 1

50: Loop Until i >= LON

End Sub

Private Sub LoadResource(ByVal mIndex As Integer)

    Dim temp() As String
    
    Dim i As Integer
    
    Dim mCount As Integer

    temp = Split(gameString(rndNum(mIndex)), ",")
      
    ReDim ptpos(UBound(temp) - 3) As Integer
    
    diffCount = (UBound(ptpos) + 1) / 4
       
    ReDim isOverSelect(diffCount - 1) As Boolean
    
    For i = 0 To UBound(ptpos)
        ptpos(i) = CInt(temp(i + 2))
    Next
     
    Picture1.Picture = LoadPicture(App.Path & "\res\" & temp(0))
    Picture2.Picture = LoadPicture(App.Path & "\res\" & temp(1))
    
    Label1.Caption = "还剩" & CStr(diffCount) & "不同"
End Sub
 

Private Sub Form_Activate()
    SetCursorPos 0.75 * 1024, 0.5 * 768
    
    ''starting from 1
    Call LoadResource(0)

End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Picture1.Left = GAME_STARTLEFT
    Picture1.Top = GAME_TOP
    Picture1.Width = GAME_WIDTH
    Picture1.Height = GAME_HEIGHT
    
    Picture2.Left = Picture1.Left + Picture1.Width + GAME_DIFF
    Picture2.Top = GAME_TOP
    Picture2.Width = GAME_WIDTH
    Picture2.Height = GAME_HEIGHT
    
    Open App.Path & "\res\pos.txt" For Input As #1
    ''////starting from 1
    Do
        i = i + 1
        ReDim Preserve gameString(i)
        Line Input #1, gameString(i)
    Loop Until EOF(1)
    
    Close #1
    
    picCount = CInt(UBound(gameString))
    
    Call DiffRndNum(picCount)
    
End Sub


Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim clkX As Integer, clkY As Integer
    Dim isFind As Boolean
   
    Static mCount As Integer
    Static nextPicture  As Integer
    
    
    
    clkX = Int(X / GAME_PUL)
    clkY = Int(Y / GAME_PUL)
    'Dim count  As Integer
     
    For t = 0 To diffCount - 1
        'Print isOverSelect(t)
        For i = ptpos(t * 4 + 0) To ptpos(t * 4 + 2)
            
            For j = ptpos(t * 4 + 1) To ptpos(t * 4 + 3)
                'count = count + 1
                If clkX = i And clkY = j Then
                   
                    If isOverSelect(t) = False Then
                        isFind = True
                        isOverSelect(t) = True
                    End If
                     
                    Exit For
                    
                End If
            Next
        Next
          
        If isFind = True Then
            
            Picture2.Line (ptpos(t * 4 + 0) * GAME_PUL, ptpos(t * 4 + 1) * GAME_PUL)-(ptpos(t * 4 + 2) * GAME_PUL, ptpos(t * 4 + 3) * GAME_PUL), vbRed, B
            Picture1.Line (ptpos(t * 4 + 0) * GAME_PUL, ptpos(t * 4 + 1) * GAME_PUL)-(ptpos(t * 4 + 2) * GAME_PUL, ptpos(t * 4 + 3) * GAME_PUL), vbRed, B
            
            mCount = mCount + 1
            Label1.Caption = "还剩 " & CStr(diffCount - mCount) & " 处不同"
            If mCount = diffCount Then '找出这张图所有不同
                mCount = 0
                nextPicture = nextPicture + 1
                Cls
                'MsgBox diffCount
                If nextPicture >= picCount Then MsgBox "你已找出所有不同":  Exit Sub '找出所有图的不同
                LoadResource (nextPicture)
            End If
            Exit For
        End If
        
    Next

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Left = X - (Picture2.Left - Picture1.Left) + Picture1.Width
    Image1.Top = Y - 12 ''local correction
    
End Sub
