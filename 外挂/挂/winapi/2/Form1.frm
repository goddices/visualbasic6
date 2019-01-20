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
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3960
      Top             =   2520
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "alpha"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "用鼠标点击问道，按回车键，将文本中的内容发给我"
      Height          =   540
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1740
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
        dx As Long
        dy As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
    ByVal hwnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long _
) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hwnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long _
) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any _
) As Long

Private Declare Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long _
)

Private Const WM_LBUTTONDOWN = &H201
Private Const WM_CHAR = &H102

Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101

Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105
Private Const WM_SYSCHAR = &H106

Private Const VK_RETURN = &HD
Private Const VK_UP = &H26

Private Const VK_MENU = &H12


Private hwndwendao As Long

Private Sub Command1_Click()


     
    hwndwendao = FindWindow("asktao", vbNullString)
     
    Text1.Text = CStr(hwndwendao)
    

End Sub

Private Sub Command2_Click()
    Print hwndwendao
    'SendMessage hwndwendao, WM_KEYDOWN, VK_UP, 0
    'PostMessage hwndwendao, WM_KEYDOWN, VK_RETURN, 0
    'PostMessage hwndwendao, WM_KEYUP, VK_RETURN, 0
    Timer1.Enabled = True
    PostMessage hwndwendao, WM_KEYDOWN, VK_MENU, &H20000001 'ALT键按下
    PostMessage hwndwendao, WM_KEYDOWN, 90, &H20000001 'E键按下必须要把第29位设置成1，代表alt键已经下
    PostMessage hwndwendao, WM_CHAR, 90, &H20000001 ' 发送一个系统字符E
    PostMessage hwndwendao, WM_KEYUP, 90, &H80000001 'E键放开，必须把31位设置成1，表示这个是系统键
    PostMessage hwndwendao, WM_KEYUP, VK_MENU, &H80000001 '

End Sub

Private Sub Timer1_Timer()

    keybd_event 18, 0, 0, 0
    keybd_event 122, 0, 0, 0
    'SendMessage dhwnd, WM_KEYDOWN, VK_UP, 0
    'PostMessage hwndwendao, WM_KEYDOWN, VK_RETURN, 0
    ''PostMessage hwndwendao, WM_KEYUP, VK_RETURN, 0
End Sub
