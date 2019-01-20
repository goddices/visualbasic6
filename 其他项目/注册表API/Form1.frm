VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const ERROR_SUCCESS = 0&

Private Const REG_SZ = 1                            ' 字符串值
Private Const REG_EXPAND_SZ = 2                    ' 可扩充字符串值
Private Const REG_BINARY = 3                        ' 二进制值
Private Const REG_DWORD = 4                        ' DWORD值
Private Const REG_MULTI_SZ = 7                     ' 多字符串值


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA " (ByVal hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA " (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

'LONG RegCreateKey(
'  HKEY hKey,        // handle to an open key
'  LPCTSTR lpSubKey, // subkey name
'  PHKEY phkResult   // buffer for key handle
');

'LONG RegOpenKey(
'  HKEY hKey,        // handle to open key
'  LPCTSTR lpSubKey, // name of subkey to open
'  PHKEY phkResult   // handle to open key
');

'这两个函数的区别在于，如果键不存在，RegCreateKey 将创建一个新键，而 RegOpenKey 将返回错误

Dim phkResult As Long  '  HKEY phkResult (void*)
Dim dwValue As Long
Dim startUpPath As String
Private Sub Command1_Click()
    If (RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", phkResult) = ERROR_SUCCESS) Then
        '// 在这里就可以使用hKey来操作daheng_directx这个KEY里面的值了。
        MsgBox "注册表操作已成功~！" & vbNewLine & Hex(phkResult)
        If RegSetValueEx(phkResult, "MYAPP1", 0, REG_SZ, ByVal CStr(startUpPath), LenB(StrConv(startUpPath, vbFromUnicode)) + 1) = ERROR_SUCCESS Then
            MsgBox "值已设定"
        End If
        
    End If
    Call RegCloseKey(phkResult)

End Sub

Private Sub Command2_Click()
    Dim strTest As String
    strTest = "alexander"
    MsgBox Hex(VarPtr(strTest))
    MsgBox Hex(StrPtr(strTest))
End Sub

' if (RegCreateKey(HKEY_LOCAL_MACHINE, "Software\daheng_directx", VarPtr(hkey)) = ERROR_SUCCESS) then
'       do something
' end if
' call regclosekey(hkey)

'LONG RegSetValueEx(
'  HKEY hKey,           // handle to key
'  LPCTSTR lpValueName, // value name , a string
'  DWORD Reserved,      // reserved    ,must be 0
'  DWORD dwType,        // value type   , REG_
'  CONST BYTE *lpData,  // value data   , various to set the value
'  DWORD cbData         // size of value data  ,sizeof(value)
');

 


Private Sub Form_Load()
    dwValue = &H7ABC
    startUpPath = Chr(34) & "c:\palette.exe" & Chr(34)
  
    
End Sub

