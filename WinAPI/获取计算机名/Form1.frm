VERSION 5.00
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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function GetSystemMenu Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal bRevert As Long _
) As Integer
Private Declare Function RemoveMenu Lib "user32" ( _
ByVal hMenu As Long, _
ByVal nPosition As Long, _
ByVal wFlags As Long _
) As Integer

Private Sub Command1_Click()
Dim Name As String, Length As Long

Length = 225
Name = String(Length, Chr(0))
GetComputerName Name, Length
Name = Left(Name, Length)
Label1.Caption = Name

End Sub

Private Sub Command2_Click()
Label2.Caption = CStr(GetTickCount())
End Sub


 

Private Sub Form_Load()
Dim R As Integer
MyMenu = GetSystemMenu(Me.hwnd, 0)
RemoveMenu MyMenu, &HF060, R
End Sub


  

