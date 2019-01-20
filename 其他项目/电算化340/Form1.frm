VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "会计电算化 选择题"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9570
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   9570
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Cmd 
      Caption         =   "退 出"
      Height          =   375
      Index           =   5
      Left            =   7800
      TabIndex        =   16
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "提 交"
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   15
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "最后个"
      Height          =   375
      Index           =   3
      Left            =   4800
      TabIndex        =   14
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "下一题"
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   13
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "上一题"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   12
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "第一题"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   11
      Top             =   7560
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   180
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   9
      Top             =   5880
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   8
      Top             =   4920
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   3960
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "题号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "您选择的答案："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   21
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "题"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5160
      TabIndex        =   20
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "转至第"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   19
      Top             =   240
      Width           =   720
   End
   Begin VB.Label LblFld 
      AutoSize        =   -1  'True
      Caption         =   "LblFld"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   7
      Left            =   2280
      TabIndex        =   18
      Top             =   6840
      Width           =   1560
   End
   Begin VB.Label LblFld 
      AutoSize        =   -1  'True
      Caption         =   "LblFld"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   6
      Left            =   5400
      TabIndex        =   17
      Top             =   6840
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label LblFld 
      AutoSize        =   -1  'True
      Caption         =   "LblFld"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Top             =   5880
      Width           =   7680
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblFld 
      AutoSize        =   -1  'True
      Caption         =   "LblFld"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   1320
      TabIndex        =   4
      Top             =   4920
      Width           =   7680
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblFld 
      AutoSize        =   -1  'True
      Caption         =   "LblFld"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   3960
      Width           =   7680
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblFld 
      AutoSize        =   -1  'True
      Caption         =   "LblFld"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   3000
      Width           =   7680
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblFld 
      Caption         =   "LblFld"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   8760
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblFld 
      AutoSize        =   -1  'True
      Caption         =   "LblFld"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim adoRs As New ADODB.Recordset
Dim temp As String
Dim sqlcheck As String
'Dim SQL As String
Dim rs As String
Dim fd(7) As ADODB.Field
Dim i As Integer
Dim store As Integer
Dim rw(1 To 340) As String
Dim c As Integer

Private Sub Cmd_Click(Index As Integer)
Select Case Index
    Case 0 '第一个
        adoRs.MoveFirst
        Ss
        FieldShow
    Case 1 '上一个
        If Not (adoRs.EOF And adoRs.BOF) Then
            adoRs.MovePrevious
            If Not adoRs.BOF Then
                FieldShow
            Else
                adoRs.MoveNext
            End If
            Ss
        End If
    Case 2 '下一个
        If Not (adoRs.EOF And adoRs.BOF) Then
            adoRs.MoveNext
            If Not adoRs.EOF Then
                FieldShow
            Else
                adoRs.MovePrevious
            End If
            Ss
        End If
     Case 3 '最后一个
        adoRs.MoveLast
        
        FieldShow
        Ss
     Case 4
        adoRs.MoveFirst
        Do While Not adoRs.EOF
            c = c + 1
            'adoRs.Fields("选择的答案")
            'adoRs!正确答案
            'adoRs!选择的答案
            
            If Trim(adoRs.Fields("正确答案")) = Trim(adoRs.Fields("选择的答案")) Then
                    store = store + 1
                    rw(c) = "第" & c & "题：正确"
                Form2.List1.AddItem rw(c)
                Else
                    rw(c) = "第" & c & "题：错误"
                    Form2.List1.AddItem rw(c)
            End If
            adoRs.MoveNext
        Loop
        adoRs.MoveFirst
        adoRs.Move Int(LblFld(0)) - 1
        Ss
        FieldShow
        Form2.Label1.Caption = "您的得分是： " & store & "分"
        Form2.Label1.FontSize = 14
        Form2.Show vbModal
        c = 0
        store = 0
        
        
    Case 5
        Unload Form2
        Unload Me
        End
End Select
End Sub

Private Sub Combo1_Click()
adoRs.MoveFirst
adoRs.Move Combo1.ListIndex
Ss
FieldShow

End Sub


Private Sub Form_Load()

Set cn = New ADODB.Connection
temp = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\340！.mdb;Jet OLEDB:Database Password= 49518"
cn.Open (temp)  '打开与数据库的连接
'If cn.State = adStateOpen Then MsgBox "Connection to NorthWind Successful!"
Set adoRs = New ADODB.Recordset
adoRs.Open "select*from Sheet1", cn, adOpenStatic, adLockOptimistic



For i = 0 To 7
    Set fd(i) = adoRs.Fields(i)
    LblFld(i).FontSize = 14
Next i
'fd8 = Format(adoRs.Fields("选择的答案"))
FieldShow

For cmbi = 1 To adoRs.RecordCount
Combo1.AddItem cmbi
Next cmbi
Combo1.Text = "1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set adoRs = cn.Execute("update Sheet1 set 选择的答案='' ")
'adoRs.Close

cn.Close
Set cn = Nothing
Set adoRs = Nothing
End
End Sub

Sub FieldShow()
'On Error Resume Next

    
For i = 1 To 6
    LblFld(i).Caption = Format(adoRs.Fields(i).Value)
Next i
    LblFld(0).Caption = fd(0)
    fd(7) = Format(adoRs.Fields("选择的答案").Value)
    LblFld(7).Caption = fd(7)
End Sub

Sub Ss()
 For i = 0 To 3
        If Option1(i).Caption = adoRs.Fields("选择的答案") Then
            Option1(i).Value = True
        Else
            Option1(i).Value = False
      
        End If
Next i
End Sub



Private Sub Option1_Click(Index As Integer)

'adoRs.Fields ("选择的答案")
adoRs.Fields("选择的答案").Value = Chr(Index + 65)
FieldShow
End Sub

