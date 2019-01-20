VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "图片处理"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12225
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   422
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   815
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command4 
      Caption         =   "清除编辑框"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "生成资源"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打开原图"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommDlg 
      Left            =   10680
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开编辑后的图片"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Height          =   5250
      Left            =   6120
      ScaleHeight     =   346
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   376
      TabIndex        =   1
      Top             =   720
      Width           =   5700
   End
   Begin VB.PictureBox Picture1 
      Height          =   5250
      Left            =   240
      ScaleHeight     =   346
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   376
      TabIndex        =   0
      Top             =   720
      Width           =   5700
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const GAME_PUL = 10
Private Const GAME_WIDTH = 380
Private Const GAME_HEIGHT = 350

Private Const ArrLenX = GAME_WIDTH / GAME_PUL
Private Const ArrLenY = GAME_HEIGHT / GAME_PUL

Private startX As Single, startY As Single

Private endX As Single, endY As Single

Private unitStartX As Integer, unitStartY  As Integer

Private unitEndX As Integer, unitEndY As Integer

Private selectRegion  As Boolean

Private strPoint As String

Private mCount As Integer

Private rdrw() As Variant

Private arr() As String

Private file1 As String
 
Private file2 As String

Private Sub FillArray(ByVal i As Integer, ByVal sx As Integer, ByVal sy As Integer, ByVal ex As Integer, ByVal ey As Integer)
    ReDim Preserve rdrw(i)
    Print i
    rdrw(i) = CStr(sx) & "," & CStr(sy) & "," & CStr(ex) & "," & CStr(ey)
End Sub

Private Sub ReDraw()
    On Error Resume Next ''第一次使用下标越界
    For i = 1 To UBound(rdrw)
        arr = Split(rdrw(i), ",")
        Picture1.Line (arr(0), arr(1))-(arr(2), arr(3)), vbRed, B
    Next
End Sub

Private Function FileName(ByVal f As String) As String
    FileName = Mid(f, InStrRev(f, "\") + 1)
End Function

Private Function OpenImageFile() As String
    Dim str As String
    CommDlg.Filter = "所有图片文件|*.jpg;*.bmp;*.gif"
    CommDlg.ShowOpen
    str = CommDlg.FileName
    OpenImageFile = str
End Function

Private Sub Command1_Click()
    file1 = OpenImageFile()
    If file1 <> "" Then Picture1.Picture = LoadPicture(file1)
End Sub

Private Sub Command2_Click()
    file2 = OpenImageFile()
    If file2 <> "" Then Picture2.Picture = LoadPicture(file2)
End Sub

Private Sub Command3_Click()
    If file1 = "" Or file2 = "" Then
        MsgBox "文件未找到！", vbExclamation, "FILE NOT FOUND"
    Else
        Open App.Path & "\res\pos.txt" For Append As #1
        Print #1, FileName(file1) & "," & FileName(file2) & "," & strPoint
        Close #1
        'MsgBox file1 & "  " & App.Path & "\res\" & Mid(file1, InStrRev(App.Path, "\") + 1)
        FileCopy file1, App.Path & "\res\" & FileName(file1)
        FileCopy file2, App.Path & "\res\" & FileName(file2)
        strPoint = ""
    End If
    Picture1.Cls
End Sub

Private Sub Command4_Click()
    Picture1.Cls
    mCount = 0
    ReDim arr(0) As String
    ReDim rdrw(0) As Variant
    strPoint = ""
End Sub

 

Private Sub Form_Load()
    Picture1.DrawWidth = 2

End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim unitX As Single, unitY As Single
    
    If Button = 1 Then
        startX = X
        startY = Y
        unitStartX = Int(X / GAME_PUL)
        unitStartY = Int(Y / GAME_PUL)
    
    ElseIf Button = 2 Then
    
        If selectRegion = True Then
        
            
            If unitStartX < 0 Then unitStartX = 0
            If unitStartX >= GAME_WIDTH Then unitStartX = GAME_WIDTH
            If unitStartY < 0 Then unitStartY = 0
            If unitStartY >= GAME_HEIGHT Then unitStartX = GAME_HEIGHT
                
            If unitEndX < 0 Then unitEndX = 0
            If unitEndX >= GAME_WIDTH Then unitEndX = GAME_WIDTH
            If unitEndY < 0 Then unitEndY = 0
            If unitEndY >= GAME_HEIGHT Then unitEndX = GAME_HEIGHT
            
            Call Swap(unitStartX, unitEndX)
            Call Swap(unitStartY, unitEndY)
            
            strPoint = strPoint & unitStartX & "," & unitStartY & "," & unitEndX & "," & unitEndY & ","

            mCount = mCount + 1
            
            Call FillArray(mCount, startX, startY, endX, endY)
            
            selectRegion = False
            
        End If
        
    End If
 
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    If Button = 1 Then
        selectRegion = True
        Picture1.Cls
        Call ReDraw
        Picture1.Line (startX, startY)-(X, Y), vbRed, B
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        endX = X
        endY = Y
        unitEndX = Int(X / GAME_PUL)
        unitEndY = Int(Y / GAME_PUL)
    End If
End Sub

Private Sub Swap(a As Integer, b As Integer)
    Dim temp As Integer
    If a > b Then
        temp = a
        a = b
        b = temp
    End If
End Sub
