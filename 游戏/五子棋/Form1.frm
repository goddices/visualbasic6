VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   1035
   ClientTop       =   1410
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   6870
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   5000
      Left            =   240
      ScaleHeight     =   4995
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   240
      Width           =   5000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const GAME_WIDTH = 15              '棋谱
Private Const DRAW_WIDTH = 120
Private a As Integer, b As Integer          '五子棋的坐标
Private x_basic As Single, y_basic As Single  '基本长度， 用于画棋子的参数
Private pcX As Single, pcY As Single        '棋子的画图坐标
Private bCtrlFlag As Boolean                '控制权
Private myArray(-5 To GAME_WIDTH + 5, -5 To GAME_WIDTH + 5) As Integer     '五子棋坐标系
Private bIsWin(6) As Boolean             '判断水平方向的胜利条件
'Private V_win(5) As Boolean             '     垂直
'Private LR_win(5) As Boolean                '左上右下
'Private RL_win(5) As Boolean                ' 左下右上
Private intWinCount As Integer

Private Sub Command1_Click()
Initialize
End Sub

Private Sub Command2_Click()
Call test
End Sub

Private Sub Form_Activate()
Initialize
End Sub

Private Sub Initialize()
    Picture1.Cls
    Picture1.AutoRedraw = True
    'Picture1.ScaleHeight = 8000
    ''Picture1.ScaleWidth = 8000
    'Picture1.Height = 8000
    'Picture1.Width = 8000
    x_basic = Picture1.Width / GAME_WIDTH
    y_basic = Picture1.Height / GAME_WIDTH
    Picture1.DrawWidth = 1
    For i = 0 To 15
        m = i * x_basic + 0.5 * x_basic
             
        Picture1.Line (m, 0.5 * y_basic)-(m, Picture1.Height - 0.5 * y_basic)
        Picture1.Line (0.5 * x_basic, m)-(Picture1.Width - 0.5 * x_basic, m)
    Next i
    
    bCtrlFlag = True
    
    For i = 0 To 6
        bIsWin(i) = False
    Next
    
    For i = -5 To GAME_WIDTH + 5
        For j = -5 To GAME_WIDTH + 5
            myArray(i, j) = 0
        Next j
    Next i

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then

    a = Int(X / x_basic)
    b = Int(Y / y_basic)
    
    'Delta_a = X / x_basic - a
    'Delta_b = Y / y_basic - b
    
    pcX = a * x_basic + 0.5 * x_basic
    pcY = b * y_basic + 0.5 * y_basic
  

    If ((myArray(a, b) = 0) And (bCtrlFlag = True)) Then
    
        Picture1.FillColor = RGB(0, 0, 0)
      
       
        myArray(a, b) = 1
        bCtrlFlag = False
        Picture1.Circle (pcX, pcY), DRAW_WIDTH
        
        For i = 1 To 4
            If FiveBoot(i, a, b, 1) = True Then
                MsgBox "Black Win"
            End If
        Next
     
    Else
    
        If ((myArray(a, b) = 0) And (bCtrlFlag = False)) Then
          Picture1.FillColor = RGB(255, 255, 255)
        
        Picture1.Circle (pcX, pcY), 120
        'Picture1.PSet (pcX, pcY), vbWhite
        myArray(a, b) = 2
        bCtrlFlag = True
        Picture1.Circle (pcX, pcY), DRAW_WIDTH
        
        For i = 1 To 4
            If FiveBoot(i, a, b, 2) = True Then
                MsgBox "White Win"
            End If
        Next

        End If
    End If
    
    
End If
End Sub

Sub test()

MsgBox bIsWin(3)
End Sub

Function FiveBoot(ByVal Direction As Byte, ByVal i As Integer, ByVal j As Integer, _
                  ByVal theColor As Integer _
) As Boolean ' Direction:H V RL RL ,i:a j:b ,theColor:Black White

    Select Case Direction
        Case 1 '// H
        
            Do While myArray(i, j) = theColor
                i = i - 1
            Loop
            
            i = i + 1
            
            For i = i To i + 5
                intWinCount = intWinCount + 1
                If myArray(i, j) = theColor Then
                    bIsWin(intWinCount) = True
                Else
                    bIsWin(intWinCount) = False
                End If
            Next i
                     
        Case 2 '//V
        
            Do While myArray(i, j) = theColor
                j = j - 1
            Loop
            
            j = j + 1
            
            For j = j To j + 5
                intWinCount = intWinCount + 1
                If myArray(i, j) = theColor Then
                    bIsWin(intWinCount) = True
                Else
                    bIsWin(intWinCount) = False
                End If
            Next j
            
        Case 3 'LR
        
            Do While myArray(i, j) = theColor
                i = i - 1
                j = j - 1
            Loop
            
            i = i + 1
            
            For i = i To i + 5
                intWinCount = intWinCount + 1
                j = j + 1
                If myArray(i, j) = theColor Then
                    bIsWin(intWinCount) = True
                Else
                    bIsWin(intWinCount) = False
                End If
            Next i
        
        Case 4 ' RL
        
            Do While myArray(i, j) = theColor
                i = i + 1
                j = j - 1
            Loop
            
            j = j + 1
            
            For j = j To j + 5
                intWinCount = intWinCount + 1
                i = i - 1
                If myArray(i, j) = theColor Then
                    bIsWin(intWinCount) = True
                Else
                    bIsWin(intWinCount) = False
                End If
            Next j
 
    End Select

    intWinCount = 0
    
    If bIsWin(1) And bIsWin(2) And bIsWin(3) And bIsWin(4) And bIsWin(5) = True And bIsWin(6) <> True Then
        FiveBoot = True
    Else
        FiveBoot = False
    End If
End Function


