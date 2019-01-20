VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   9375
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdrestart 
      Caption         =   "Command2"
      Height          =   255
      Left            =   7080
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Command1"
      Height          =   255
      Left            =   7080
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blackturn As Boolean '�ֵ��ڷ�����
Dim whiteturn As Boolean '�ֵ��׷�����
Dim table(0 To 15, 0 To 15) As Integer '�ô˶�ά�����ʾ����
Dim inti As Integer '����Ԫ�ء���
Dim intj As Integer
Dim boolstatus As Boolean '��ʾ���״̬������/����

Private Sub cmdclose_Click() '�رմ���
Unload Me
Set frmmain = Nothing
End Sub

Private Sub cmdrestart_Click() '���¿�ʼ
'�������
Me.Cls

'��������
For inti = 0 To 15
 For intj = 0 To 15
 table(inti, intj) = 0
 Next
Next

'�ػ�����
Form_Load
End Sub

Private Sub Form_Load()
'������
Form_Paint
blackturn = True '�ڷ�����
boolstatus = True '��ʼ
Label1.Caption = "�ڷ�����"
End Sub

'����
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim intx As Integer '���Ӻ���λ��
Dim inty As Integer '��������λ��

'ȷ������Ƿ��ڽ����У�������
If boolstatus = False Then
 Label1.Caption = "����"
 Exit Sub
End If

'ȷ�����ӵ�ȷ��λ��
'��������λ�ò��������У�������
If x < 10 Or x > 310 Or y < 10 Or y > 310 Then
 Exit Sub
End If
'��������λ���������У���ת��Ϊ��Ӧ�������ӵ������
If (x - 10) Mod 20 < 10 Then
 intx = x - (x - 10) Mod 20
Else
 intx = x + 20 - (x - 10) Mod 20
End If
If (y - 10) Mod 20 < 10 Then
 inty = y - (y - 10) Mod 20
Else
 inty = y + 20 - (y - 10) Mod 20
End If

'������ת���������е���ӦԪ��
inti = (intx - 10) / 20
intj = (inty - 10) / 20

'���������Ԫ�ز�Ϊ�㣬����ʾ��������Ӧ���������ӣ�������
If table(inti, intj) <> 0 Then
 Exit Sub
End If

'���ӣ�Բ��
If blackturn = True Then
 '��ɫ
 Me.FillColor = RGB(0, 0, 0)
 table(inti, intj) = 1 '���Ӹ�1
 Label1.Caption = "�׷�"
Else
 '��ɫ
 Me.FillColor = RGB(255, 255, 255)
 table(inti, intj) = 2 '���Ӹ�2
 Label1.Caption = "�ڷ�"
End If
Me.FillStyle = 0 '����ȱ
Me.Circle (intx, inty), 8

'�ж��Ƿ�����������
Call judgeman

'����
blackturn = Not blackturn 'ȡ��


End Sub
Private Sub judgeman() '�ж��Ƿ�����������

Dim strwho As String '���ӷ�����

If table(inti, intj) = 1 Then '��ʾ�ڷ��µ���
 strwho = "�ڷ�"
Else
 strwho = "�׷�"
End If

'�ֱ��жϺ������Խ����Ƿ������ӣ��˶δ���Ƚϸ��ӣ�����������⣬����ִ��Ч�ʼ���
'�ǳ��ʺ������̸��Ӻܶ�����

If samelinenums(1, 0) >= 5 Or samelinenums(0, 1) >= 5 Or samelinenums(1, 1) >= 5 Or samelinenums(-1, 1) >= 5 Then
 MsgBox strwho & "ʤ��"
 boolstatus = False '��ֽ���
End If
End Sub

Function samelinenums(changei As Integer, changej As Integer) '�ж�ͬһֱ���ϵ�������
Dim i As Integer
Dim j As Integer
Dim num As Integer 'ͬһ������ͬ��ɫ������

'��������һ��ͬ��ɫ��������
i = inti: j = intj
Do
 If table(i, j) <> table(inti, intj) Then
 num = max(Abs(inti - i), Abs(intj - j))
 Exit Do
 End If
 i = i + changei: j = j + changej
Loop Until i < 0 Or i > 15 Or j < 0 Or j > 15

'����������һ��ͬ��ɫ��������
i = inti: j = intj
Do
 If table(i, j) <> table(inti, intj) Then
 num = num - 1 + max(Abs(inti - i), Abs(intj - j))
 Exit Do
 End If
 i = i - changei: j = j - changej
Loop Until i < 0 Or i > 15 Or j < 0 Or j > 15
'MsgBox num
samelinenums = num
End Function

'��ϴ�ֵ
Function max(inta As Integer, intb As Integer)
 max = inta
 If max < intb Then max = intb
End Function

Private Sub Form_Paint() '��(10,10)Ϊ���Ͻ����껭һ��16*16,ÿ��߳�Ϊ20���ص�����
Cls '���
Dim i As Integer
ScaleMode = 3 '�趨���廭���ĵ�λΪ����
For i = 10 To 330 Step 20
 Me.Line (10, i)-(330, i)
 Me.Line (i, 10)-(i, 330)
Next
End Sub

