VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
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
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Menu dd 
      Caption         =   "asdf"
      Visible         =   0   'False
      Begin VB.Menu sadf 
         Caption         =   "asdf"
      End
      Begin VB.Menu ddd 
         Caption         =   "12"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'void Shell_NotifyIconExample()
'{
'    // This code will add a Shell_NotifyIcon notificaion on PocketPC and Smartphone
'    NOTIFYICONDATA nid = {0};
'    nid.cbSize = sizeof(nid);
'    nid.uID = 100;      // Per WinCE SDK docs, values from 0 to 12 are reserved and should not be used.
'    nid.uFlags = NIF_ICON;
'    nid.hIcon = LoadIcon(g_hInstance, MAKEINTRESOURCE(IDI_SAMPLEICON));

'    //Add the notification to the tray
'    Shell_NotifyIcon(NIM_ADD, &nid);

'    //Update the icon of the notification
'    nid.uFlags = NIF_ICON;
'    nid.hIcon = LoadIcon(g_hInstance, MAKEINTRESOURCE(IDI_SAMPLEICON2));
'    Shell_NotifyIcon(NIM_MODIFY, &nid);

'    //remove the notification from the tray
 '   Shell_NotifyIcon(NIM_DELETE, &nid);

'    return;
'}


Private Sub Command1_Click()
'Shell_NothifyIcon_Example
Dim aa As Long
aa = dd()
Print aa
End Sub

Private Sub Command2_Click()
Print GetWindowLong(Me.hwnd, GWL_WNDPROC)
End Sub


