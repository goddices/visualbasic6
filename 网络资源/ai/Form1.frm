VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   6225
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txtq 
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3720
      Width           =   5535
   End
   Begin VB.CommandButton cmdtalk 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtans 
      Height          =   3255
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'COPYRIGHT scѰ
'VER 0.01
'��Ҫ˵��
'��������һ���򵥵����������
'ʹ�������������������iqq()�в�����û�а�������Ĺؼ��ʣ������>=1 ���������ѡһ��ش�
'���û�����rnda()�������ѡһ������𰸻ش�
'���������cmdtalk��ť txtans txtq �ı��� multiline��Ϊtrue scrollbars��Ϊ2
Dim iqa() As String 'iq�ش�ĵ�����
Dim iqq() As String 'iq�ش�ĵ�������
Dim iqnum As Integer 'iq�ش�����⡢�𰸸���
Dim rnda() As String ' rnd�ش�ĵ�����
Dim rndansnum As Integer 'rnd�ش�Ĵ𰸸���
Dim dat As String '���ļ������ԭʼ����
Dim username As String 'ʹ��������
Dim robotname As String '����������
Dim ver As String '���ݰ汾
Dim rndans As String 'rnd�ش����������
Dim iqans As String 'iq�ش����������



Private Sub cmdtalk_Click()

Dim q As String
Dim a As String
q = txtq.Text
txtans.Text = txtans.Text & username & ":" & vbNewLine
txtans.Text = txtans.Text & q & vbNewLine & vbNewLine
'���ȼ��iqq

Dim manyiq As Integer

For i = 1 To iqnum
If InStr(q, iqq(i)) <> 0 Then
manyiq = manyiq + 1
End If
Next i

ReDim manyans(manyiq) As String
Dim oneiq As Integer

For i = 1 To iqnum
If InStr(q, iqq(i)) <> 0 Then
oneiq = oneiq + 1
manyans(oneiq) = iqa(i)
End If
Next i

If manyiq <> 0 Then

a = manyans(Int(Rnd * oneiq) + 1)
Else

a = rnda(Int(Rnd * rndansnum) + 1)
End If
txtans.Text = txtans.Text & robotname & ":" & vbNewLine
txtans.Text = txtans.Text & a & vbNewLine & vbNewLine
txtq.SetFocus
txtans.SelStart = Len(txtans.Text)
End Sub

Private Sub Form_Load()
Randomize
datname = App.Path + "\talk.dat"
Open datname For Binary As 1
dat = Space(LOF(1))
Get 1, , dat
Close 1

'��ȡusername robotname ver rndans iqans
username = InputBox("�������������", "��ʾ")
start = InStr(dat, "<name>")
over = InStr(dat, "</name>")
robotname = Mid(dat, start + 6, over - start - 6)
start = InStr(dat, "<ver>")
over = InStr(dat, "</ver>")
ver = Mid(dat, start + 5, over - start - 5)
start = InStr(dat, "<rndans>")
over = InStr(dat, "</rndans>")
rndans = Mid(dat, start + 8, over - start - 8)
start = InStr(dat, "<iqans>")
over = InStr(dat, "</iqans>")
iqans = Mid(dat, start + 7, over - start - 7)

'rnda(rndansnum)��ȡ

Dim rndanslen As Long 'rnd�ش���������ݳ���
rndansnum = 0
rndanslen = Len(rndans)
For i = 1 To rndanslen
If Mid(rndans, i, 1) = "|" Then
rndansnum = rndansnum + 1
End If
Next i

ReDim rnda(rndansnum) As String '����ÿһ��rnd�ش�Ķ�̬����
Dim lastl As Integer '��һ��|��λ��
num = 1 'num�������еı�ţ�rndansnum��������
For i = 1 To rndanslen
If Mid(rndans, i, 1) = "|" Then
rnda(num) = Mid(rndans, lastl + 1, i - lastl - 1)
num = num + 1
lastl = i
End If
Next i

'iqq(rndansnum) iqa(rndansnum)��ȡ

Dim iqanslen As Long
iqanslen = Len(iqans)
For i = 1 To iqanslen
If Mid(iqans, i, 1) = "|" Then
iqansnum = iqansnum + 1
End If
Next i

ReDim iqa(iqansnum) As String
ReDim iqq(iqansnum) As String

lasta = 0
lastq = 0
num = 1
For i = 1 To iqanslen
If Mid(iqans, i, 1) = "\" Then
iqq(num) = Mid(iqans, lasta + 1, i - lasta - 1)
num = num + 1
lastq = i
ElseIf Mid(iqans, i, 1) = "|" Then
iqa(num - 1) = Mid(iqans, lastq + 1, i - lastq - 1)
lasta = i
End If
Next i

iqnum = iqansnum

'��ӭ��Ϣ

txtans.Text = txtans.Text & username & ",��ӭʹ�ñ����" _
& vbNewLine & "�ҽ�" & robotname & " ���ݿ�汾 " & ver & vbNewLine
txtans.Text = txtans.Text & "iqnum " & iqnum & vbNewLine
txtans.Text = txtans.Text & "rndansnum " & rndansnum & vbNewLine & vbNewLine
End Sub


Private Sub txtq_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call cmdtalk_Click
End Sub
