VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   960
      Top             =   2520
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Text1.Text = LCase(Text1.Text)
Do While Not (Rs.BOF And Rs.EOF)
    If Text1.Text = Rs.Fields("UserName") Then
        If Text2.Text = Rs.Fields("PWD") Then
            MsgBox "welcome  " & Text1.Text
            Form2.Show
        Else
            MsgBox "password incorrect"
         
        End If
       
        Exit Do
        
    End If

    Rs.MoveNext
    If (Rs.EOF) Then Rs.MoveFirst: Exit Do
Loop
End Sub

Private Sub Form_Load()
Call OpenDateBase
End Sub

Private Sub Form_Unload(Cancel As Integer)
Rs.Close
Set Rs = Nothing
Conn.Close
Set Conn = Nothing
End Sub

