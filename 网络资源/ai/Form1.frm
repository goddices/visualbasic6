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
   StartUpPosition =   3  '窗口缺省
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
'COPYRIGHT sc寻
'VER 0.01
'简要说明
'本程序是一个简单的聊天机器人
'使用者提出的问题首先在iqq()中查找有没有包含定义的关键词，如果有>=1 个则随机挑选一句回答
'如果没有则从rnda()中随机挑选一句随机答案回答
'窗体中添加cmdtalk按钮 txtans txtq 文本框 multiline设为true scrollbars设为2
Dim iqa() As String 'iq回答的单个答案
Dim iqq() As String 'iq回答的单个问题
Dim iqnum As Integer 'iq回答的问题、答案个数
Dim rnda() As String ' rnd回答的单个答案
Dim rndansnum As Integer 'rnd回答的答案个数
Dim dat As String '从文件载入的原始数据
Dim username As String '使用者名字
Dim robotname As String '机器人名字
Dim ver As String '数据版本
Dim rndans As String 'rnd回答的所有数据
Dim iqans As String 'iq回答的所有数据



Private Sub cmdtalk_Click()

Dim q As String
Dim a As String
q = txtq.Text
txtans.Text = txtans.Text & username & ":" & vbNewLine
txtans.Text = txtans.Text & q & vbNewLine & vbNewLine
'首先检查iqq

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

'获取username robotname ver rndans iqans
username = InputBox("请输入你的名字", "提示")
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

'rnda(rndansnum)获取

Dim rndanslen As Long 'rnd回答的所有数据长度
rndansnum = 0
rndanslen = Len(rndans)
For i = 1 To rndanslen
If Mid(rndans, i, 1) = "|" Then
rndansnum = rndansnum + 1
End If
Next i

ReDim rnda(rndansnum) As String '定义每一个rnd回答的动态数组
Dim lastl As Integer '上一次|的位置
num = 1 'num是数组中的标号（rndansnum是总数）
For i = 1 To rndanslen
If Mid(rndans, i, 1) = "|" Then
rnda(num) = Mid(rndans, lastl + 1, i - lastl - 1)
num = num + 1
lastl = i
End If
Next i

'iqq(rndansnum) iqa(rndansnum)获取

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

'欢迎信息

txtans.Text = txtans.Text & username & ",欢迎使用本软件" _
& vbNewLine & "我叫" & robotname & " 数据库版本 " & ver & vbNewLine
txtans.Text = txtans.Text & "iqnum " & iqnum & vbNewLine
txtans.Text = txtans.Text & "rndansnum " & rndansnum & vbNewLine & vbNewLine
End Sub


Private Sub txtq_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call cmdtalk_Click
End Sub
