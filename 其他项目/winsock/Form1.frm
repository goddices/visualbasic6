VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   3930
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3000
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const BASE64CHR         As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
Private psBase64Chr(0 To 63)    As String



Private Sub Command1_Click()
    Winsock1.Connect   '������Զ�̼���������ӡ�
End Sub

 

Private Sub Form_Load()
    Winsock1.LocalPort = 0       '���ñ���ʹ�õĶ˿�
    Winsock1.Protocol = sckTCPProtocol       '����Winsock�ؼ�ʹ�õ�Э�飬TCP��UDP��
    Winsock1.RemoteHost = "smtp.163.com"  '���÷���Email�ķ�����
    Winsock1.RemotePort = 25       '����Ҫ���ӵ�Զ�̶˿ں�
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
    First = "mail from:" + Chr(32) + "azftestmail@163.com" + vbCrLf                   '�����˵�ַ
    Snd = "rcpt to:" + Chr(32) + "470857673@qq.com " + vbCrLf                   '�����˵�ַ
    DateNow = Format(Date, "Ddd ") & ",   " & Format(Date, "dd Mmm YYYY ") & "   " & Format(Time, "hh:mm:ss ") & " " & "   -0600 "
    Third = "date:" + Chr(32) + DateNow + vbCrLf                   '��ʼ����ʱ��
    Fourth = "From:" + Chr(32) + "azftestmail@163.com" + vbCrLf                   '����������
    Fifth = "To:" + Chr(32) + "alexanderzhufeng" + vbCrLf                '����������
    Sixth = "Subject:" + Chr(32) + "VB   С԰����֪ͨ " + vbCrLf                 '���ŵ�����
    Seventh = "VB   С԰�Ѿ����� " + vbCrLf           '���ŵ�����
    Ninth = "X-Mailer: EBT Reporter v 2.x" + vbCrLf
    Eighth = Fourth + Third + Ninth + Fifth + Sixth
    
    Winsock1.SendData ("EHLO smtp.163.com" + vbCrLf)          '��ʼ����
    Winsock1.SendData ("AUTH LOGIN" + vbCrLf)
    
    Winsock1.SendData (Base64("azftestmail") + vbCrLf)
    Winsock1.SendData (Base64("19891025") + vbCrLf)
    
    Winsock1.SendData ("mail from:" + Chr(32) + "azftestmail@163.com" + vbCrLf)
    Winsock1.SendData ("rcpt to:" + Chr(32) + "470857673@qq.com " + vbCrLf)
    Winsock1.SendData ("data" + vbCrLf)
    Winsock1.SendData ("Date:" + Chr(32) + Format(Date, "Ddd") & "," & Format(Date, "dd Mmm YYYY") & "" & Format(Time, "hh:mm:ss") & "" & "-0600" + vbCrLf)
    Winsock1.SendData ("From:" + Chr(32) + "xiaopeng" + vbCrLf)
    Winsock1.SendData ("X-Mailer: vbemailsender" + vbCrLf)
    Winsock1.SendData ("To:" + Chr(32) + "lingling" + vbCrLf)
    Winsock1.SendData ("Subject:" + Chr(32) + "how are you" + vbCrLf)
    Winsock1.SendData ("ni hao ma" + vbCrLf)
    Winsock1.SendData ("." + vbCrLf)
    Winsock1.SendData ("quit " + vbCrLf)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next         '�ڴ������������󣬻ָ�ԭ�е�����
    Dim webData     As String
    Winsock1.GetData webData, vbString       'ȡ�÷��ź�ķ�����Ϣ�����Լ���Ƿ����
    Text1.Text = Text1.Text + webData

End Sub



Private Function Base64(ByVal Str As String) As String 'base6�����㷨
    Const BASE64_TABLE As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim StrTempLine As String
    Dim j As Integer
    For j = 1 To (Len(Str) - Len(Str) Mod 3) Step 3
        StrTempLine = StrTempLine + Mid(BASE64_TABLE, (Asc(Mid(Str, j, 1)) \ 4) + 1, 1)
        StrTempLine = StrTempLine + Mid(BASE64_TABLE, ((Asc(Mid(Str, j, 1)) Mod 4) * 16 _
                      + Asc(Mid(Str, j + 1, 1)) \ 16) + 1, 1)
        StrTempLine = StrTempLine + Mid(BASE64_TABLE, ((Asc(Mid(Str, j + 1, 1)) Mod 16) * 4 _
                      + Asc(Mid(Str, j + 2, 1)) \ 64) + 1, 1)
        StrTempLine = StrTempLine + Mid(BASE64_TABLE, (Asc(Mid(Str, j + 2, 1)) Mod 64) + 1, 1)
    Next j
    If Not (Len(Str) Mod 3) = 0 Then
        If (Len(Str) Mod 3) = 2 Then
            StrTempLine = StrTempLine + Mid(BASE64_TABLE, (Asc(Mid(Str, j, 1)) \ 4) + 1, 1)
            StrTempLine = StrTempLine + Mid(BASE64_TABLE, (Asc(Mid(Str, j, 1)) Mod 4) * 16 _
                      + Asc(Mid(Str, j + 1, 1)) \ 16 + 1, 1)
            StrTempLine = StrTempLine + Mid(BASE64_TABLE, (Asc(Mid(Str, j + 1, 1)) Mod 16) * 4 + 1, 1)
            StrTempLine = StrTempLine & "="
        ElseIf (Len(Str) Mod 3) = 1 Then
            StrTempLine = StrTempLine + Mid(BASE64_TABLE, Asc(Mid(Str, j, 1)) \ 4 + 1, 1)
            StrTempLine = StrTempLine + Mid(BASE64_TABLE, (Asc(Mid(Str, j, 1)) Mod 4) * 16 + 1, 1)
            StrTempLine = StrTempLine & "=="
        End If
    End If
    Base64 = StrTempLine
End Function


