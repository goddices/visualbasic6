VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "&Get"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   2880
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    '
    Dim lngRetVal      As Long
    Dim strErrorMsg    As String
    Dim udtWinsockData As WSAData
    Dim lngType        As Long
    Dim lngProtocol    As Long
    '
    'start up winsock service
    lngRetVal = WSAStartup(&H101, udtWinsockData)
    '
    If lngRetVal <> 0 Then
        '
        '
        Select Case lngRetVal
            Case WSASYSNOTREADY
                strErrorMsg = "The underlying network subsystem is not " & _
                    "ready for network communication."
            Case WSAVERNOTSUPPORTED
                strErrorMsg = "The version of Windows Sockets API support " & _
                    "requested is not provided by this particular " & _
                    "Windows Sockets implementation."
            Case WSAEINVAL
                strErrorMsg = "The Windows Sockets version specified by the " & _
                    "application is not supported by this DLL."
        End Select
        '
        MsgBox strErrorMsg, vbCritical
        '
    End If
    '
End Sub
 
 Private Sub Form_Unload(Cancel As Integer)
    Call WSACleanup
End Sub

Private Sub ShowErrorMsg(lngError As Long)
    '
    Dim strMessage As String
    '
    Select Case lngError
        Case WSANOTINITIALISED
            strMessage = "A successful WSAStartup call must occur " & _
                         "before using this function."
        Case WSAENETDOWN
            strMessage = "The network subsystem has failed."
        Case WSAHOST_NOT_FOUND
            strMessage = "Authoritative answer host not found."
        Case WSATRY_AGAIN
            strMessage = "Nonauthoritative host not found, or server failure."
        Case WSANO_RECOVERY
            strMessage = "A nonrecoverable error occurred."
        Case WSANO_DATA
            strMessage = "Valid name, no data record of requested type."
        Case WSAEINPROGRESS
            strMessage = "A blocking Windows Sockets 1.1 call is in " & _
                         "progress, or the service provider is still " & _
                         "processing a callback function."
        Case WSAEFAULT
            strMessage = "The name parameter is not a valid part of " & _
                         "the user address space."
        Case WSAEINTR
            strMessage = "A blocking Windows Socket 1.1 call was " & _
                         "canceled through WSACancelBlockingCall."
    End Select
    '
    MsgBox strMessage, vbExclamation
    '
End Sub


Private Sub cmdExit_Click()
    Unload Me
End Sub


