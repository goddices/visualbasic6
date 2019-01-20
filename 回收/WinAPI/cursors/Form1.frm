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
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
End
End Sub

Private Sub Form_Activate()

Call GetMessage(uMsg, Form1.hwnd, 0, 0)

Select Case uMsg.message

    Case WM_LBUTTONDOWN
    
        MessageBox Form1.hwnd, CStr(Hex(uMsg.message)), "dfdfdf", MB_OK + MB_ICONINFORMATION
    
 
    
    Case WM_DESTROY
        End
        'PostQuitMessage (0)

        
End Select

 
 
End
End Sub

