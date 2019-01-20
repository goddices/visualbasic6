Attribute VB_Name = "Module1"
Public Declare _
Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
    ByVal idHook As Long, _
    ByVal lpfn As Long, _
    ByVal hmod As Long, _
    ByVal dwThreadId As Long _
) As Long

Public Declare Function UnhookWindowsHookEx Lib "user32" ( _
    ByVal hHook As Long _
) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long _
)

Public Declare Function CallNextHookEx Lib "user32" ( _
    ByVal hHook As Long, _
    ByVal ncode As Long, _
    ByVal wParam As Long, _
    lParam As Any _
) As Long

 
'typedef struct {
'    DWORD vkCode;
'    DWORD scanCode;
'    DWORD flags;
'    DWORD time;
'    ULONG_PTR dwExtraInfo;
'} KBDLLHOOKSTRUCT, *PKBDLLHOOKSTRUCT;

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtranInfo As Long
End Type



Public Const WH_CALLWNDPROC = 4
Public Const WH_JOURNALPLAYBACK = 1
Public Const WH_MOUSE = 7
Public Const WH_KEYBOARD = 2
Public Const WH_JOURNALRECORD = 0
Public Const WH_KEYBOARD_LL = 13 '#define WH_KEYBOARD_LL     13
 
Public Const WM_KEYDOWN = &H100

Public hHook As Long
Public strREC As String
Public kbs As KBDLLHOOKSTRUCT


Sub Main()
    hHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf HookFunc, App.hInstance, 0)
    Form1.Show
     
End Sub

Public Function HookFunc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Call CopyMemory(ByVal VarPtr(kbs), ByVal lParam, Len(kbs))
    If wParam = WM_KEYDOWN Then
        Form1.Label1.Caption = "VK code : " & kbs.vkCode
        strREC = strREC + CStr(Chr(kbs.vkCode))
        
        
    End If
    HookFunc = CallNextHookEx(hHook, ncode, wParam, lParam)
End Function

