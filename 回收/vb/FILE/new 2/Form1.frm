VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommDlg1 
      Left            =   840
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   720
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "open"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim FileStr As String
    Dim str As String
    Dim h As String
    Dim ar As Byte
    CommDlg1.Filter = "所图片文件|*.jpg;*.bmp;*.gif"
    CommDlg1.ShowOpen
    
    FileStr = CommDlg1.FileName
    If FileStr <> "" Then
        Open FileStr For Binary As #1
        For i = LOF(1) - 100 To LOF(1)
            Get #1, i, ar
             
            str = str & ar & vbNewLine
        Next
        
        Close #1
    End If
    
    Dim strarr() As String
    strarr = Split(str, vbNewLine)
    For i = 0 To UBound(strarr)
       Text1.Text = Text1.Text + vbNewLine + strarr(i)
    Next
    
     
End Sub
