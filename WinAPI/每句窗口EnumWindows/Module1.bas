Attribute VB_Name = "Module1"
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lparam As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Const STR_asktao As String = "AskTao"
'
 

Public nIndex As Long
 
Public hWD() As Long

Public Function ProcFunc(ByVal hwnd As Long, ByVal lparam As Long) As Long
  
    Dim strClsName As String
    strClsName = Space(255)
    If hwnd <> 0 Then
         
        'ReDim Preserve strClsName(nIndex) As String
        GetClassName hwnd, strClsName, 255
        
        If Left(strClsName, Len("AskTao")) = "AskTao" Then
            nIndex = nIndex + 1
            ReDim Preserve hWD(nIndex) As Long
            hWD(nIndex) = hwnd
            SetWindowText hwnd, "SB" & CStr(nIndex) & "ºÅ"
            ProcFunc = 1
        End If
        ProcFunc = 1
    Else
        ProcFunc = 0
    End If
End Function

Public Sub wocao()
    Call EnumWindows(AddressOf ProcFunc, 0)
End Sub
