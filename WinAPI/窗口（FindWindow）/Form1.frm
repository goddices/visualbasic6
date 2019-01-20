VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Handle to Window"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   11160
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame Frame4 
      Caption         =   "by handle to window ( hwnd )"
      Height          =   3135
      Left            =   5640
      TabIndex        =   23
      Top             =   2640
      Width           =   5295
      Begin VB.TextBox TxtWndTxt4 
         Height          =   375
         Left            =   1680
         TabIndex        =   27
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox TxtClsName4 
         Height          =   375
         Left            =   1680
         TabIndex        =   26
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CommandButton Command4 
         Caption         =   "execute"
         Height          =   375
         Left            =   3600
         TabIndex        =   25
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "its wnd text"
         Height          =   180
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "its class name"
         Height          =   180
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   1260
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "by pointer position"
      Height          =   3135
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   5295
      Begin VB.CommandButton Command3 
         Caption         =   "execute"
         Height          =   375
         Left            =   3600
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtHWnd3 
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox TxtClsName3 
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox TxtWndTxt3 
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   2400
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "its hwnd"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "its class name"
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1260
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "its wnd text"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   1080
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "by class name"
      Height          =   2295
      Left            =   5640
      TabIndex        =   7
      Top             =   240
      Width           =   5295
      Begin VB.CommandButton Command2 
         Caption         =   "execute"
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox TxtHWnd2 
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox TxtWndTxt2 
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "its hwnd"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   720
      End
      Begin VB.Label Label9 
         Caption         =   "its wnd text"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "by window text"
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      Begin VB.CommandButton Command1 
         Caption         =   "execute"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox TxtClsName1 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox TxtHWnd1 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "its class name"
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "its hwnd"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   720
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
        px As Long
        py As Long
End Type

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
    ByVal hwnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long _
) As Long


Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long _
) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hwnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long _
) As Long

Private Declare Function WindowFromPoint Lib "user32" ( _
    ByVal xPoint As Long, _
    ByVal yPoint As Long _
) As Long

Private Declare Function GetCursorPos Lib "user32" ( _
    lpPoint As POINTAPI _
) As Long

Private str As String
Private Const STR_LEN = 255

Private Sub Command1_Click()
Dim hwindow As Long
If Text1.Text <> "" Then
    hwindow = FindWindow(vbNullString, Text1.Text)
    GetClassName hwindow, str, STR_LEN
    TxtHWnd1.Text = CStr(hwindow)
    TxtClsName1.Text = str
End If
End Sub

Private Sub Command2_Click()
Dim hwindow As Long
If Text2.Text <> "" Then
    hwindow = FindWindow(Text2.Text, vbNullString)
    GetWindowText hwindow, str, STR_LEN
    TxtHWnd2.Text = CStr(hwindow)
    TxtWndTxt2.Text = str
End If
End Sub

Private Sub Command3_Click()
Dim hwindow As Long
Dim pt As POINTAPI
    GetCursorPos pt
    hwindow = WindowFromPoint(pt.px, pt.py)
    TxtHWnd3.Text = CStr(hwindow)
    GetClassName hwindow, str, STR_LEN
    TxtClsName3.Text = str
    GetWindowText hwindow, str, STR_LEN
    TxtWndTxt3.Text = str
 
End Sub

Private Sub Command4_Click()
Dim hwindow As Long
If Text4.Text <> 0 And IsNumeric(Text4.Text) Then
    hwindow = CLng(Text4.Text)
    GetClassName hwindow, str, STR_LEN
    TxtClsName4.Text = str
    GetWindowText hwindow, str, STR_LEN
    TxtWndTxt4.Text = str

End If
End Sub

Private Sub Form_Load()
str = String(STR_LEN, 0)
End Sub

Private Sub Text3_Click()
Command3.SetFocus
End Sub
