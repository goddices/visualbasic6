VERSION 5.00
Begin VB.Form frmForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "旋转俄罗斯 1.0 Demo -- 泰立软件工作室"
   ClientHeight    =   6345
   ClientLeft      =   1275
   ClientTop       =   705
   ClientWidth     =   4950
   Icon            =   "frmForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6345
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraFrameNext 
      Caption         =   "下一块"
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "隐藏(&D)"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1095
      End
      Begin VB.PictureBox picPictureNextBackGround 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1275
         ScaleWidth      =   1050
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   1110
         Begin VB.Image imgPictureNext 
            Height          =   495
            Left            =   120
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   480
      Top             =   4800
   End
   Begin VB.PictureBox picBackGround 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   6030
      Left            =   1680
      ScaleHeight     =   20
      ScaleMode       =   0  'User
      ScaleWidth      =   10
      TabIndex        =   0
      Top             =   120
      Width           =   3030
      Begin VB.PictureBox picPictureTemp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1680
         ScaleHeight     =   480
         ScaleWidth      =   1080
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.PictureBox picPictureNow 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   480
         ScaleHeight     =   495
         ScaleWidth      =   975
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "泰立软件工作室荣誉出品                    作者：尹强"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Image imgPictureNowBackup 
      Height          =   375
      Left            =   960
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuGame 
      Caption         =   "游戏(&G)"
      Begin VB.Menu mnuGameNew 
         Caption         =   "新游戏(&N)"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpKey 
         Caption         =   "键盘(&K)"
      End
      Begin VB.Menu mnuGameAbout 
         Caption         =   "关于(&A)"
      End
   End
End
Attribute VB_Name = "frmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Type_Now As Integer '目前方块的类型
Dim Type_Next As Integer '下个方块的类型
Dim intRotate As Integer '方块旋转的状态

Function Get_X_Value()
  If GetValue(1, 2) Then   'Get X Value
    If MaxX - MinX >= 2 Then
      If MaxX - CurX <= 1 Then
        Adjust_Left = MaxX - 2 - 1
      Else
        Adjust_Left = CurX - 1
      End If
      Get_X_Value = True
      Exit Function
    End If
  End If
  Get_X_Value = False
End Function

Function GetValue(nType As Integer, nWid As Integer)
    GetCoor
    On Error Resume Next
    Dim OKCount, EmptyCount As Integer
    MinX = Xs(1).cX
    MaxX = Xs(1).cX
    MinY = Xs(1).cY
    MaxY = Xs(1).cY
    For i = 2 To 4
        If MinX > Xs(i).cX Then MinX = Xs(i).cX
        If MaxX < Xs(i).cX Then MaxX = Xs(i).cX
        If MinY > Xs(i).cY Then MinY = Xs(i).cY
        If MaxY < Xs(i).cY Then MaxY = Xs(i).cY
    Next
    For i = MinX To MaxX
        For j = MinY To MaxY
            If Total(i, j) Then
                GetValue = False
                Exit Function
            End If
        Next
    Next
                
                If nType = 0 Then   'Get Y Value
                    EmptyCount = 0  'Get MinY
                    OKCount = 0
                    For i = MinY - 1 To MinY - (nWid - 1) Step -1
                        
                        For j = MinX To MaxX
                            If Total(j, i) = False Then OKCount = OKCount + 1
                        Next
                        If OKCount >= picPictureNow.Width And OKCount >= picPictureNow.Height Then
                            EmptyCount = EmptyCount + 1
                            OKCount = 0
                        Else
                            Exit For
                        End If
                    Next
                    MinY = MinY - EmptyCount
                    If MinY < 1 Then MinY = 1
                    
                    EmptyCount = 0  'GetMaxY
                    OKCount = 0
                    For i = MaxY + 1 To MaxY + nWid - 1
                        For j = MinX To MaxX
                            If Total(j, i) = False Then OKCount = OKCount + 1
                        Next
                        If OKCount >= picPictureNow.Width And OKCount >= picPictureNow.Height Then
                            EmptyCount = EmptyCount + 1
                            OKCount = 0
                        Else
                            Exit For
                        End If
                    Next
                    MaxY = MaxY + EmptyCount
                    If MaxY > 20 Then MaxY = 20
                    
                Else    'Get X Value
                    EmptyCount = 0  'Get MinX
                    OKCount = 0
                    For i = MinX - 1 To MinX - (nWid - 1) Step -1
                        
                        For j = MinY To MaxY
                            If Total(i, j) = False Then OKCount = OKCount + 1
                        Next
                        If OKCount >= picPictureNow.Width And OKCount >= picPictureNow.Height Then
                            EmptyCount = EmptyCount + 1
                            OKCount = 0
                        Else
                            Exit For
                        End If
                    Next
                    MinX = MinX - EmptyCount
                    If MinX < 1 Then MinX = 1
                    
                    EmptyCount = 0  'GetMaxX
                    OKCount = 0
                    For i = MaxX + 1 To MaxX + (nWid - 1)
                        For j = MinY To MaxY
                            If Total(i, j) = False Then OKCount = OKCount + 1
                        Next
                        If OKCount >= picPictureNow.Width And OKCount >= picPictureNow.Height Then
                            EmptyCount = EmptyCount + 1
                            OKCount = 0
                        Else
                            Exit For
                        End If
                    Next
                    MaxX = MaxX + EmptyCount
                    If MaxX > 10 Then MaxX = 10
                End If
    GetValue = True
End Function

Function Get_Y_Value()
                    If GetValue(0, 2) Then    'Get Y Value
                        If MaxY - MinY >= 2 Then
                            If MaxY - (picPictureNow.Top + 1) <= 1 Then
                                Adjust_Top = MinY - 1
                            Else
                                Adjust_Top = picPictureNow.Top
                            End If
                            Get_Y_Value = True
                            Exit Function
                        End If
                    End If
                    Get_Y_Value = False
End Function

Sub Global_Init()
'全局初始化
picBackGround.Cls
imgPictureNext.Picture = LoadPicture("")
picPictureNow.Visible = False
tmrDrop.Enabled = False
End Sub

Sub Init()
'每个方块的初始化过程
picPictureNow.Visible = False
tmrDrop.Enabled = False
Type_Now = Type_Next
picPictureNow.Picture = imgPictureNext.Picture
imgPictureNowBackup.Picture = picPictureNow.Picture
Sel_Next
intRotate = 0
picPictureNow.Left = 4
picPictureNow.Top = 0
picPictureNow.Visible = True
tmrDrop.Enabled = True
End Sub

Sub GetCoor()
'获取一个方块的 4 个点的坐标
For i = 1 To 4  'init
    Xs(i).cX = 0
    Xs(i).cY = 0
    Xs(i).cZ = False
Next
CurX = picPictureNow.Left + 1
        Select Case Type_Now
            Case 1  '长条
                If intRotate Mod 2 = 1 Then
                    Xs(1).cX = CurX
                    Xs(1).cY = picPictureNow.Top + 1
                    Xs(1).cZ = True
                    For i = 2 To 4
                        Xs(i).cX = CurX + i - 1
                        Xs(i).cY = picPictureNow.Top + 1
                        Xs(i).cZ = True
                    Next
                Else
                    Xs(1).cX = CurX
                    Xs(1).cY = picPictureNow.Top + 4
                    Xs(1).cZ = True
                    For i = 2 To 4
                        Xs(i).cX = CurX
                        Xs(i).cY = picPictureNow.Top + i - 1
                        Xs(i).cZ = False
                    Next
                End If
            Case 2  '2字
                If intRotate Mod 2 = 1 Then
                    Xs(1).cX = CurX
                    Xs(1).cY = picPictureNow.Top + 3
                    Xs(1).cZ = True
                    Xs(2).cX = CurX + 1
                    Xs(2).cY = picPictureNow.Top + 2
                    Xs(2).cZ = True
                    For i = 3 To 4
                        Xs(i).cX = CurX + i - 3
                        Xs(i).cY = picPictureNow.Top + 5 - i
                        Xs(i).cZ = False
                    Next
                Else
                    Xs(1).cX = CurX
                    Xs(1).cY = picPictureNow.Top + 1
                    Xs(1).cZ = True
                    Xs(2).cX = CurX + 1
                    Xs(2).cY = picPictureNow.Top + 2
                    Xs(2).cZ = True
                    Xs(3).cX = CurX + 2
                    Xs(3).cY = picPictureNow.Top + 2
                    Xs(3).cZ = True
                    Xs(4).cX = CurX + 1
                    Xs(4).cY = picPictureNow.Top + 1
                    Xs(4).cZ = False
                End If
            Case 3  '7字
                Select Case intRotate Mod 4
                    Case 0
                        Xs(1).cX = CurX
                        Xs(1).cY = picPictureNow.Top + 1
                        Xs(1).cZ = True
                        Xs(2).cX = CurX + 1
                        Xs(2).cY = picPictureNow.Top + 3
                        Xs(2).cZ = True
                        For i = 3 To 4
                            Xs(i).cX = CurX + 1
                            Xs(i).cY = picPictureNow.Top + i - 2
                            Xs(i).cZ = False
                        Next
                    Case 1
                        Xs(1).cX = CurX
                        Xs(1).cY = picPictureNow.Top + 2
                        Xs(1).cZ = True
                        Xs(2).cX = CurX + 1
                        Xs(2).cY = picPictureNow.Top + 2
                        Xs(2).cZ = True
                        Xs(3).cX = CurX + 2
                        Xs(3).cY = picPictureNow.Top + 2
                        Xs(3).cZ = True
                        Xs(4).cX = CurX + 2
                        Xs(4).cY = picPictureNow.Top + 1
                        Xs(4).cZ = False
                    Case 2
                        Xs(1).cX = CurX
                        Xs(1).cY = picPictureNow.Top + 3
                        Xs(1).cZ = True
                        Xs(2).cX = CurX + 1
                        Xs(2).cY = picPictureNow.Top + 3
                        Xs(2).cZ = True
                        For i = 3 To 4
                            Xs(i).cX = CurX
                            Xs(i).cY = picPictureNow.Top + i - 2
                            Xs(i).cZ = False
                        Next
                    Case 3
                        Xs(1).cX = CurX
                        Xs(1).cY = picPictureNow.Top + 2
                        Xs(1).cZ = True
                        Xs(2).cX = CurX + 1
                        Xs(2).cY = picPictureNow.Top + 1
                        Xs(2).cZ = True
                        Xs(3).cX = CurX + 2
                        Xs(3).cY = picPictureNow.Top + 1
                        Xs(3).cZ = True
                        Xs(4).cX = CurX
                        Xs(4).cY = picPictureNow.Top + 1
                        Xs(4).cZ = False
                End Select
            Case 4  'T字
                Select Case intRotate Mod 4
                    Case 0
                        Xs(1).cX = CurX
                        Xs(1).cY = picPictureNow.Top + 2
                        Xs(1).cZ = True
                        Xs(2).cX = CurX + 1
                        Xs(2).cY = picPictureNow.Top + 2
                        Xs(2).cZ = True
                        Xs(3).cX = CurX + 2
                        Xs(3).cY = picPictureNow.Top + 2
                        Xs(3).cZ = True
                        Xs(4).cX = CurX + 1
                        Xs(4).cY = picPictureNow.Top + 1
                        Xs(4).cZ = False
                    Case 1
                        Xs(1).cX = CurX
                        Xs(1).cY = picPictureNow.Top + 3
                        Xs(1).cZ = True
                        Xs(2).cX = CurX + 1
                        Xs(2).cY = picPictureNow.Top + 2
                        Xs(2).cZ = True
                        For i = 3 To 4
                            Xs(i).cX = CurX
                            Xs(i).cY = picPictureNow.Top + i - 2
                            Xs(i).cZ = False
                        Next
                    Case 2
                        Xs(1).cX = CurX
                        Xs(1).cY = picPictureNow.Top + 1
                        Xs(1).cZ = True
                        Xs(2).cX = CurX + 1
                        Xs(2).cY = picPictureNow.Top + 2
                        Xs(2).cZ = True
                        Xs(3).cX = CurX + 2
                        Xs(3).cY = picPictureNow.Top + 1
                        Xs(3).cZ = True
                        Xs(4).cX = CurX + 1
                        Xs(4).cY = picPictureNow.Top + 1
                        Xs(4).cZ = False
                    Case 3
                        Xs(1).cX = CurX
                        Xs(1).cY = picPictureNow.Top + 2
                        Xs(1).cZ = True
                        Xs(2).cX = CurX + 1
                        Xs(2).cY = picPictureNow.Top + 3
                        Xs(2).cZ = True
                        For i = 3 To 4
                            Xs(i).cX = CurX + 1
                            Xs(i).cY = picPictureNow.Top + i - 2
                            Xs(i).cZ = False
                        Next
                End Select
            Case 5  '反7字
                Select Case intRotate Mod 4
                    Case 0
                        Xs(1).cX = CurX
                        Xs(1).cY = picPictureNow.Top + 3
                        Xs(1).cZ = True
                        Xs(2).cX = CurX + 1
                        Xs(2).cY = picPictureNow.Top + 1
                        Xs(2).cZ = True
                        For i = 3 To 4
                            Xs(i).cX = CurX
                            Xs(i).cY = picPictureNow.Top + i - 2
                            Xs(i).cZ = False
                        Next
                    Case 1
                        Xs(1).cX = CurX
                        Xs(1).cY = picPictureNow.Top + 1
                        Xs(1).cZ = True
                        Xs(2).cX = CurX + 1
                        Xs(2).cY = picPictureNow.Top + 1
                        Xs(2).cZ = True
                        Xs(3).cX = CurX + 2
                        Xs(3).cY = picPictureNow.Top + 2
                        Xs(3).cZ = True
                        Xs(4).cX = CurX + 2
                        Xs(4).cY = picPictureNow.Top + 1
                        Xs(4).cZ = False
                    Case 2
                        Xs(1).cX = CurX
                        Xs(1).cY = picPictureNow.Top + 3
                        Xs(1).cZ = True
                        Xs(2).cX = CurX + 1
                        Xs(2).cY = picPictureNow.Top + 3
                        Xs(2).cZ = True
                        For i = 3 To 4
                            Xs(i).cX = CurX + 1
                            Xs(i).cY = picPictureNow.Top + i - 2
                            Xs(i).cZ = False
                        Next
                    Case 3
                        Xs(1).cX = CurX
                        Xs(1).cY = picPictureNow.Top + 2
                        Xs(1).cZ = True
                        Xs(2).cX = CurX + 1
                        Xs(2).cY = picPictureNow.Top + 2
                        Xs(2).cZ = True
                        Xs(3).cX = CurX + 2
                        Xs(3).cY = picPictureNow.Top + 2
                        Xs(3).cZ = True
                        Xs(4).cX = CurX
                        Xs(4).cY = picPictureNow.Top + 1
                        Xs(4).cZ = False
                End Select
            Case 6  '反2字
                If intRotate Mod 2 = 1 Then
                    Xs(1).cX = CurX
                    Xs(1).cY = picPictureNow.Top + 2
                    Xs(1).cZ = True
                    Xs(2).cX = CurX + 1
                    Xs(2).cY = picPictureNow.Top + 3
                    Xs(2).cZ = True
                    For i = 3 To 4
                        Xs(i).cX = CurX + i - 3
                        Xs(i).cY = picPictureNow.Top + i - 2
                        Xs(i).cZ = False
                    Next
                Else
                    Xs(1).cX = CurX
                    Xs(1).cY = picPictureNow.Top + 2
                    Xs(1).cZ = True
                    Xs(2).cX = CurX + 1
                    Xs(2).cY = picPictureNow.Top + 2
                    Xs(2).cZ = True
                    Xs(3).cX = CurX + 2
                    Xs(3).cY = picPictureNow.Top + 1
                    Xs(3).cZ = True
                    Xs(4).cX = CurX + 1
                    Xs(4).cY = picPictureNow.Top + 1
                    Xs(4).cZ = False
                End If
            Case 7  '田字
                    Xs(1).cX = CurX
                    Xs(1).cY = picPictureNow.Top + 2
                    Xs(1).cZ = True
                    Xs(2).cX = CurX + 1
                    Xs(2).cY = picPictureNow.Top + 2
                    Xs(2).cZ = True
                    For i = 3 To 4
                        Xs(i).cX = CurX + i - 3
                        Xs(i).cY = picPictureNow.Top + 1
                        Xs(i).cZ = False
                    Next
        End Select
End Sub

Sub Judge_Full()
'判断是否堆满
R_Value = picPictureNow.Top + 1   'MinY
rx_value = picPictureNow.Top + picPictureNow.Height 'MaxY
For i = rx_value To R_Value Step -1
    If Total(1, i) And Total(2, i) And Total(3, i) And Total(4, i) And Total(5, i) And _
      Total(6, i) And Total(7, i) And Total(8, i) And Total(9, i) And Total(10, i) Then
            '如果一行已经堆满，则将此行上面的图象全部向下移动一点
            k = BitBlt(picBackGround.hDC, 0, 20, 200, (i - 1) * 20, picBackGround.hDC, 0, 0, vbSrcCopy)
            For j = i To 1 Step -1
                For k = 1 To 10
                    Total(k, j) = Total(k, j - 1)
                Next k
            Next j
            i = i + 1
    End If
Next i
'如果目前方块的顶点位置 <=0 ，则表示全部堆满
If picPictureNow.Top <= 0 Then
    Select Case MsgBox("你玩完了！想再试试身手吗？", 4 + 32)
      Case vbYes
        mnuGameNew_Click
      Case Else
        Global_Init
    End Select
End If
End Sub

'判断方块能否翻转
Function Judge_Rotate()
        Select Case Type_Now
            Case 1  '长条
                If intRotate Mod 2 = 1 Then
                    If GetValue(0, 4) Then    'Get Y Value
                        If MaxY - MinY >= 3 Then
                            Adjust_Top = MinY - 1
                            Judge_Rotate = True
                            Exit Function
                        End If
                    End If
                    Judge_Rotate = False
                    Exit Function
                Else
                    If GetValue(1, 4) Then 'Get X Value
                        If MaxX - MinX >= 3 Then
                            If MaxX - CurX <= 2 Then
                                Adjust_Left = MaxX - 3 - 1
                            Else
                                If CurX = MinX Then
                                    Adjust_Left = CurX - 1
                                Else
                                    Adjust_Left = CurX - 1 - 1
                                End If
                            End If
                            Judge_Rotate = True
                            Exit Function
                        End If
                    End If
                    Judge_Rotate = False
                    Exit Function
                End If
            Case 2  '2字
                If intRotate Mod 2 = 0 Then
                    Judge_Rotate = Get_Y_Value
                    Exit Function
                Else
                    Judge_Rotate = Get_X_Value
                    Exit Function
                End If
            Case 3  '7字
                Select Case intRotate Mod 4
                    Case 0
                        Judge_Rotate = Get_X_Value
                        Exit Function
                    Case 1
                        Judge_Rotate = Get_Y_Value
                        Exit Function
                    Case 2
                        Judge_Rotate = Get_X_Value
                        Exit Function
                    Case 3
                        Judge_Rotate = Get_Y_Value
                        Exit Function
                End Select
            Case 4  'T字
                Select Case intRotate Mod 4
                    Case 0
                        Judge_Rotate = Get_Y_Value
                        Exit Function
                    Case 1
                        Judge_Rotate = Get_X_Value
                        Exit Function
                    Case 2
                        Judge_Rotate = Get_Y_Value
                        Exit Function
                    Case 3
                        Judge_Rotate = Get_X_Value
                        Exit Function
                End Select
            Case 5  '反7字
                Select Case intRotate Mod 4
                    Case 0
                        Judge_Rotate = Get_X_Value
                        Exit Function
                    Case 1
                        Judge_Rotate = Get_Y_Value
                        Exit Function
                    Case 2
                        Judge_Rotate = Get_X_Value
                        Exit Function
                    Case 3
                        Judge_Rotate = Get_Y_Value
                        Exit Function
                End Select
            Case 6  '反2字
                If intRotate Mod 2 = 0 Then
                    Judge_Rotate = Get_Y_Value
                    Exit Function
                Else
                    Judge_Rotate = Get_X_Value
                    Exit Function
                End If
        End Select
End Function
'判断能否向左移动
Function JudgeX_Left()
GetCoor
For i = 1 To 4
        On Error Resume Next
        If Xs(i).cY > 0 Then
            If Total(Xs(i).cX - 1, Xs(i).cY) Or Xs(i).cX = 0 Then
                JudgeX_Left = False
                Exit Function
            End If
        End If
Next
JudgeX_Left = True
End Function
'判断能否向右移动
Function JudgeX_Right()
GetCoor
For i = 1 To 4
        On Error Resume Next
        If Xs(i).cY > 0 Then
            If Total(Xs(i).cX + 1, Xs(i).cY) Or Xs(i).cX = 10 Then
                JudgeX_Right = False
                Exit Function
            End If
        End If
Next
JudgeX_Right = True
End Function
'判断能否向下移动
Sub JudgeY()
GetCoor
For i = 1 To 4
    If Xs(i).cZ Then
        On Error Resume Next
        If Xs(i).cY > 0 Then
            If Total(Xs(i).cX, Xs(i).cY + 1) Or Xs(i).cY = 20 Then
                '如果不能移动，将4点位置的坐标设置为 True,并将图形固定下来
                For j = 1 To 4
                    Total(Xs(j).cX, Xs(j).cY) = True
                Next j
                picBackGround.PaintPicture picPictureNow.Picture, picPictureNow.Left, picPictureNow.Top, picPictureNow.Width, picPictureNow.Height, , , , , vbSrcAnd
                Judge_Full
                If picPictureNow.Visible Then Init
                Exit Sub
            End If
        End If
    End If
Next
End Sub

Sub Sel_Next()
'随机从 7 个放块中选择一个
Randomize
Type_Next = Int((7 * Rnd) + 1)
Select Case Type_Next
    Case 1
        imgPictureNext.Picture = LoadResPicture(11, 0)
    Case 2
        imgPictureNext.Picture = LoadResPicture(13, 0)
    Case 3
        imgPictureNext.Picture = LoadResPicture(15, 0)
    Case 4
        imgPictureNext.Picture = LoadResPicture(19, 0)
    Case 5
        imgPictureNext.Picture = LoadResPicture(23, 0)
    Case 6
        imgPictureNext.Picture = LoadResPicture(27, 0)
    Case 7
        imgPictureNext.Picture = LoadResPicture(29, 0)
End Select
imgPictureNext.Move (picPictureNextBackGround.Width - imgPictureNext.Width) \ 2 - 30, (picPictureNextBackGround.Height - imgPictureNext.Height) \ 2 - 30
End Sub

Private Sub cmdDisplay_Click()
imgPictureNext.Visible = Not (imgPictureNext.Visible)
If imgPictureNext.Visible Then
    cmdDisplay.Caption = "隐藏(&D)"
Else
    cmdDisplay.Caption = "显示(&S)"
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'改变 Case 的 KeyCode 值就可以改变键盘控制按钮
Select Case KeyCode
    Case vbKeyLeft
        If picPictureNow.Left - 1 >= 0 Then
            J_Value = JudgeX_Left
            If J_Value Then
                picPictureNow.Picture = imgPictureNowBackup.Picture
                r = BitBlt(picPictureTemp.hDC, 0, 0, picPictureNow.Width * 20, picPictureNow.Height * 20, picBackGround.hDC, (picPictureNow.Left - 1) * 20, picPictureNow.Top * 20, vbSrcCopy)
                picPictureNow.Left = picPictureNow.Left - 1
                r = BitBlt(picPictureNow.hDC, 0, 0, picPictureNow.Width * 20, picPictureNow.Height * 20, picPictureTemp.hDC, 0, 0, vbSrcAnd)
            End If
        End If
    Case vbKeyRight
        If picPictureNow.Left + picPictureNow.Width < picBackGround.ScaleWidth Then
            J_Value = JudgeX_Right
            If J_Value Then
                picPictureNow.Picture = imgPictureNowBackup.Picture
                r = BitBlt(picPictureTemp.hDC, 0, 0, picPictureNow.Width * 20, picPictureNow.Height * 20, picBackGround.hDC, (picPictureNow.Left + 1) * 20, picPictureNow.Top * 20, vbSrcCopy)
                picPictureNow.Left = picPictureNow.Left + 1
                r = BitBlt(picPictureNow.hDC, 0, 0, picPictureNow.Width * 20, picPictureNow.Height * 20, picPictureTemp.hDC, 0, 0, vbSrcAnd)
            End If
        End If
    Case vbKeyDown
        tmrDrop_Timer
    Case vbKeyUp
      If Judge_Rotate Then
        intRotate = intRotate + 1
        Select Case Type_Now
            Case 1  '长条
                If intRotate Mod 2 = 1 Then
                    picPictureNow.Picture = LoadResPicture(12, 0)
                    picPictureNow.Top = picPictureNow.Top + 3
                    picPictureNow.Left = Adjust_Left
                Else
                    picPictureNow.Picture = LoadResPicture(11, 0)
                    picPictureNow.Top = Adjust_Top
                    picPictureNow.Left = picPictureNow.Left + 1
                End If
            Case 2  '2字
                If intRotate Mod 2 = 1 Then
                    picPictureNow.Picture = LoadResPicture(14, 0)
                    picPictureNow.Top = Adjust_Top
                Else
                    picPictureNow.Picture = LoadResPicture(13, 0)
                    picPictureNow.Top = picPictureNow.Top + 1
                    picPictureNow.Left = Adjust_Left
                End If
            Case 3  '7字
                Select Case intRotate Mod 4
                    Case 0
                        picPictureNow.Picture = LoadResPicture(15, 0)
                        picPictureNow.Top = Adjust_Top
                    Case 1
                        picPictureNow.Picture = LoadResPicture(16, 0)
                        picPictureNow.Top = picPictureNow.Top + 1
                        picPictureNow.Left = Adjust_Left
                    Case 2
                        picPictureNow.Picture = LoadResPicture(17, 0)
                        picPictureNow.Top = Adjust_Top
                    Case 3
                        picPictureNow.Picture = LoadResPicture(18, 0)
                        picPictureNow.Top = picPictureNow.Top + 1
                        picPictureNow.Left = Adjust_Left
                End Select
            Case 4  'T字
                Select Case intRotate Mod 4
                    Case 0
                        picPictureNow.Picture = LoadResPicture(19, 0)
                        picPictureNow.Top = picPictureNow.Top + 1
                        picPictureNow.Left = Adjust_Left
                    Case 1
                        picPictureNow.Picture = LoadResPicture(20, 0)
                        picPictureNow.Top = Adjust_Top
                    Case 2
                        picPictureNow.Picture = LoadResPicture(21, 0)
                        picPictureNow.Top = picPictureNow.Top + 1
                        picPictureNow.Left = Adjust_Left
                    Case 3
                        picPictureNow.Picture = LoadResPicture(22, 0)
                        picPictureNow.Top = Adjust_Top
                End Select
            Case 5  '反7字
                Select Case intRotate Mod 4
                    Case 0
                        picPictureNow.Picture = LoadResPicture(23, 0)
                        picPictureNow.Top = Adjust_Top
                    Case 1
                        picPictureNow.Picture = LoadResPicture(24, 0)
                        picPictureNow.Top = picPictureNow.Top + 1
                        picPictureNow.Left = Adjust_Left
                    Case 2
                        picPictureNow.Picture = LoadResPicture(25, 0)
                        picPictureNow.Top = Adjust_Top
                    Case 3
                        picPictureNow.Picture = LoadResPicture(26, 0)
                        picPictureNow.Top = picPictureNow.Top + 1
                        picPictureNow.Left = Adjust_Left
                End Select
            Case 6  '反2字
                If intRotate Mod 2 = 1 Then
                    picPictureNow.Picture = LoadResPicture(28, 0)
                    picPictureNow.Top = Adjust_Top
                Else
                    picPictureNow.Picture = LoadResPicture(27, 0)
                    picPictureNow.Top = picPictureNow.Top + 1
                    picPictureNow.Left = Adjust_Left
                End If
        End Select
        imgPictureNowBackup.Picture = picPictureNow.Picture
      End If
End Select
End Sub

Private Sub mnuGameAbout_Click()
MsgBox "旋转俄罗斯 1.0 Demo", vbInformation
End Sub

Private Sub mnuGameExit_Click()
End
End Sub

Private Sub mnuGameNew_Click()
'将 10x20 的坐标全部设置为空
For i = 1 To 10
    For j = 0 To 20
        Total(i, j) = False
    Next j
Next i
CurX = 0
picBackGround.Cls
'改变 tmrDrop 的 Interval 值即可改变游戏速度
tmrDrop.Interval = 1000
Sel_Next
Init
End Sub

Private Sub mnuHelpKey_Click()
MsgBox "键盘控制方法：" + vbCrLf + "← 控制方块向左移动；" _
        + vbCrLf + "→ 控制方块向右移动；" _
        + vbCrLf + "↓ 控制方块向下快速移动；" _
        + vbCrLf + "↑ 控制方块的顺时针方向的翻转。", 64, "旋转俄罗斯 1.0 键盘操作帮助"
End Sub

Private Sub tmrDrop_Timer()
'方块下落
JudgeY
picPictureNow.Picture = imgPictureNowBackup.Picture
r = BitBlt(picPictureTemp.hDC, 0, 0, picPictureNow.Width * 20, picPictureNow.Height * 20, picBackGround.hDC, picPictureNow.Left * 20, (picPictureNow.Top + 1) * 20, vbSrcCopy)
picPictureNow.Top = picPictureNow.Top + 1
r = BitBlt(picPictureNow.hDC, 0, 0, picPictureNow.Width * 20, picPictureNow.Height * 20, picPictureTemp.hDC, 0, 0, vbSrcAnd)
DoEvents
If picPictureNow.Top + picPictureNow.Height > picBackGround.ScaleHeight Then Init
End Sub

