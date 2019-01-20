VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "测试"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   657
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1021
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      Caption         =   "操作"
      Height          =   1815
      Left            =   480
      TabIndex        =   2
      Top             =   7680
      Width           =   3255
      Begin VB.CommandButton CmdExit 
         Caption         =   "退出"
         Height          =   495
         Left            =   1800
         TabIndex        =   37
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton CmdSav 
         Caption         =   "保存提交"
         Height          =   495
         Left            =   240
         TabIndex        =   36
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton CmdNext 
         Caption         =   "下一个"
         Height          =   495
         Left            =   1800
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton CmdPre 
         Caption         =   "上一个"
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame FrmCh 
      Caption         =   "选项"
      Height          =   3135
      Left            =   360
      TabIndex        =   1
      Top             =   4200
      Width           =   14535
      Begin VB.OptionButton Option1 
         Caption         =   "D"
         Height          =   300
         Index           =   3
         Left            =   240
         TabIndex        =   41
         Top             =   2280
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "C"
         Height          =   300
         Index           =   2
         Left            =   240
         TabIndex        =   40
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "B"
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   39
         Top             =   1080
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "A"
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "D"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   47
         Top             =   2280
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "C"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   46
         Top             =   1680
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "B"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   45
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "A"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   44
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   3
         Left            =   960
         TabIndex        =   34
         Top             =   2280
         Width           =   13335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   33
         Top             =   1680
         Width           =   13335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   32
         Top             =   1080
         Width           =   13335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   31
         Top             =   480
         Width           =   13335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "第 1 题"
      Height          =   3855
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   14535
      Begin VB.CommandButton CmdMark 
         BackColor       =   &H008080FF&
         Caption         =   "标记"
         Height          =   495
         Left            =   12840
         TabIndex        =   5
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Txt1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   360
         Width           =   14055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "选择题号"
      Height          =   1815
      Left            =   3960
      TabIndex        =   3
      Top             =   7680
      Width           =   10935
      Begin VB.CommandButton CmdMulti 
         Caption         =   "多进入选题"
         Height          =   615
         Left            =   5640
         TabIndex        =   43
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   24
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   23
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   22
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   21
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   20
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   19
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   18
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   17
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   16
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   15
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   14
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   13
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   12
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   11
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   10
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   9
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   8
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   7
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   6
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   5
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   4
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   3
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   2
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   1
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   615
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim Conn As New ADODB.Connection
Dim Rs As New ADODB.Recordset

 
Dim choiced(25) As String

Dim multichoiced(4) As String

Dim pointer As Integer

Private rndNum()  As Integer
Dim rw(1 To 25) As String
Dim sta   As Integer
Dim mark(24) As Boolean


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




Private Sub CmdCh_Click()
sta = 2
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdMark_Click()
   If CmdMark.Caption = "标记" Then
        Command2(pointer).BackColor = &H8080FF
        CmdMark.Caption = "取消标记"
        mark(pointer) = True
        
    ElseIf CmdMark.Caption = "取消标记" Then
        Command2(pointer).BackColor = &H8000000F
        CmdMark.Caption = "标记"
        mark(pointer) = False
        
    End If
    ShowProcess
End Sub


Private Sub CmdMulti_Click()
If (MsgBox("做完单选后才能进行多选，您做完单选了吗？", vbYesNo + vbDefaultButton2 + vbInformation, "选择题")) = vbYes Then
CmdMulti.Enabled = False
 
sta = 1
 
Rs.Close

Call DiffRndNum(15)
temp = " where id=" & rndNum(0)
    For i = 1 To 4
        temp = temp & " or id=" & rndNum(i)
    Next
   ' MsgBox temp
    Rs.Open "select * from multi" & temp & " order by rnd(id)", Conn, 1, 1
 pointer = 0
 ShowFields

SwitchSngMlt (False)
End If
End Sub

Private Sub CmdNext_Click()
   If Not (Rs.BOF And Rs.EOF) Then
            Rs.MoveNext
            
            If Not Rs.EOF Then
                pointer = pointer + 1
                ShowFields
                ShowMulti
            Else
                Rs.MovePrevious
            End If
            
            ShowChoiced
            
        End If
        
        ShowProcess
End Sub

Private Sub CmdPre_Click()
 If Not (Rs.EOF And Rs.BOF) Then
            Rs.MovePrevious
            If Not Rs.BOF Then
                pointer = pointer - 1
                ShowFields
                ShowMulti
            Else
                Rs.MoveNext
            End If
          
            ShowChoiced
            
        End If
        ShowProcess
End Sub

Private Sub CmdSav_Click()
    Dim t
    If sta = 1 Then
                t = multichoiced
            Else
                t = choiced
            End If
        Rs.MoveFirst
        Do While Not Rs.EOF
            c = c + 1
            'rs.Fields("选择的答案")
            'rs!正确答案
            'rs!选择的答案
            
            
            If Trim(Rs.Fields("正确答案")) = t(c - 1) Then
                    store = store + 1
                    rw(c) = "第" & c & "题：正确"
                Form2.List1.AddItem rw(c)
                Else
                    rw(c) = "第" & c & "题：错误"
                    Form2.List1.AddItem rw(c)
            End If
            Rs.MoveNext
        Loop
        Rs.MoveFirst
        Rs.Move pointer
        ShowChoiced
        ShowFields
        Form2.Label1.Caption = "您的得分是： " & store & "分"
        Form2.Label1.FontSize = 14
        Form2.Show vbModal
        c = 0
        store = 0
        ShowProcess
End Sub

Private Sub Command3_Click()
 
End Sub

Private Sub Command1_Click()
Dim str As String
For i = 0 To 4
    str = str & multichoiced(i) & vbNewLine
Next
MsgBox str
End Sub

Private Sub Command2_Click(Index As Integer)
   pointer = Index
    Rs.Move Index, adBookmarkFirst
    ShowFields
    ShowChoiced
     ShowProcess
   
End Sub

Private Sub Form_Load()
     
    Call DiffRndNum(75)
    
    Dim temp As String
    Set Conn = New ADODB.Connection
    temp = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.mdb;Jet OLEDB"
    Conn.Open (temp)
    Set Rs = New ADODB.Recordset
    
    temp = " where id=" & rndNum(0)
    For i = 1 To 24
        temp = temp & " or id=" & rndNum(i)
    Next
    temp = temp & " order by rnd(ID)"
    Rs.Open "select * from choice " & temp, Conn, 1, 1
    ShowFields
   
   For i = 1 To 25
    Command2(i - 1).Caption = i
   Next
   
   CmdMulti.Enabled = True
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Rs.Close
    Set Rs = Nothing
    Conn.Close
    Set Conn = Nothing
End Sub

Private Sub ShowFields()

    Frame1.Caption = "第 " & pointer + 1 & " 题"
    Txt1.Text = Rs("题干")
    For i = 2 To 5
        Label1(i - 2).Caption = CStr(Rs(i).Value)
    Next
End Sub

 
Private Sub Option1_Click(Index As Integer)
    choiced(pointer) = Chr(Index + 65)
End Sub

Private Sub Check1_Click(Index As Integer)
 
 If Check1(Index).Value = 1 Then
        multichoiced(pointer) = multichoiced(pointer) + Chr(65 + Index)
    
 Else
    multichoiced(pointer) = Replace(multichoiced(pointer), Check1(Index).Caption, "")
 End If
'Print "选择的:" & multichoiced(pointer)
End Sub

Private Sub ShowChoiced()
Frame1.Caption = "第 " & pointer + 1 & " 题"
 For i = 0 To 3
        If Option1(i).Caption = choiced(pointer) Then
            Option1(i).Value = True
        Else
            Option1(i).Value = False
      
        End If
Next i
End Sub

Private Sub ShowMulti()
If sta = 1 Then
Dim t As String
 
   ' Cls
    Frame1.Caption = "第 " & pointer + 1 & " 题"
    'Print "选择的:" & multichoiced(pointer)
    
       For i = 0 To 3
        Check1(i).Value = 0
    Next


    For i = 1 To 4
 
          t = Mid(multichoiced(pointer), i)
          t = Left(t, 1)
            
          If t = "A" Then Check1(0).Value = 1
            If t = "B" Then Check1(1).Value = 1
           If t = "C" Then Check1(2).Value = 1
          If t = "D" Then Check1(3).Value = 1
      
      Next
     End If
End Sub

Private Sub ShowProcess()
    For i = 0 To 24
         If choiced(i) <> "" And Command2(i).BackColor <> &H8080FF Then Command2(i).BackColor = &H80FF80
        
    Next
   
    If mark(pointer) Then
            CmdMark.Caption = "取消标记"
        Else
            CmdMark.Caption = "标记"
        End If
End Sub

Private Sub SwitchSngMlt(SngOnOff As Boolean)
For i = 0 To 3
    Option1(i).Visible = SngOnOff
    Check1(i).Visible = Not SngOnOff
Next
End Sub

 

