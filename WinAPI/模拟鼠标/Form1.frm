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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

  '-------------------------------------------
  '               ģ����������������Ҽ�����
  '-------------------------------------------
  '                       �������   ��֪����
  '-------------------------------------------
  '����˵����
  '��������API����ʵ��ģ�������¼����������к��
  '����ʮ����Ȥ��Ч����Ҳ����һ�ԡ�
  '������ֻʹ�������������꣬����Ҳ����ʹ�þ���
  '�����������һ�ԡ�
  '-------------------------------------------
    
  '��VB������
  '     Private   Declare   Sub   mouse_event   Lib   "user32"   (ByVal   dwFlags   As   Long,   ByVal   dx   As   Long,   ByVal   dy   As   Long,   ByVal   cButtons   As   Long,   ByVal   dwExtraInfo   As   Long)
    
  '��˵����
  '     ģ��һ������¼�
    
  '����ע��
  '     ��������˶���ʱ����SystemParametersInfo�����涨��ϵͳ���켣�ٶȻ�Ӧ����������е��ٶ�
    
  '��������
  '     dwFlags   --------     Long��������־��һ�����
  '     MOUSEEVENTF_ABSOLUTE
  '     dx��dyָ���������ϵͳ�е�һ������λ�á����������ϵͳ�У���Ļ��ˮƽ�ʹ�ֱ�����Ͼ��ȷָ��65535��65535����Ԫ   -
  '     MOUSEEVENTF_MOVE                   �ƶ����
  '     MOUSEEVENTF_LEFTDOWN           ģ������������
  '     MOUSEEVENTF_LEFTUP               ģ��������̧��
  '     MOUSEEVENTF_RIGHTDOWN         ģ������Ҽ�����
  '     MOUSEEVENTF_RIGHTUP             ģ������Ҽ�̧��
  '     MOUSEEVENTF_MIDDLEDOWN       ģ������м�����
  '     MOUSEEVENTF_MIDDLEUP           ģ������м�̧��
  '     dx   -------------     Long�������Ƿ�ָ����MOUSEEVENTF_ABSOLUTE��־��ָ��ˮƽ����ľ���λ�û�����˶�'
    
  '     dy   -------------     Long�������Ƿ�ָ����MOUSEEVENTF_ABSOLUTE��־��ָ����ֱ����ľ���λ�û�����˶�
    
  '     cButtons   -------     Long��δʹ��
    
  '     dwExtraInfo   ----     Long��ͨ��δ�õ�һ��ֵ����GetMessageExtraInfo������ȡ�����ֵ�����õ�ֵȡ�����ض�����������
  Option Explicit
          Private Declare Sub mouse_event Lib "user32" _
          ( _
          ByVal dwFlags As Long, _
          ByVal dx As Long, _
          ByVal dy As Long, _
          ByVal cButtons As Long, _
          ByVal dwExtraInfo As Long _
          )
    
  'Option_Tag��ʾѡ������һ��ģ���¼�
  Dim Option_Tag     As Integer
  'OnTest��ʾ�Ƿ���ģ��״̬���Ա�����ֹͣģ��
  Dim OnTest     As Boolean
  '��API�����Ķ���
  Const MOUSEEVENTF_LEFTDOWN = &H2
  Const MOUSEEVENTF_LEFTUP = &H4
  Const MOUSEEVENTF_MIDDLEDOWN = &H20
  Const MOUSEEVENTF_MIDDLEUP = &H40
  Const MOUSEEVENTF_MOVE = &H1
  Const MOUSEEVENTF_ABSOLUTE = &H8000
  Const MOUSEEVENTF_RIGHTDOWN = &H8
  Const MOUSEEVENTF_RIGHTUP = &H10
    
  '����   ģ��Ŀ�ʼ�����
  Private Sub Command1_Click()
    
  '���������ģ��״̬
  If OnTest = False Then
  Command1.Caption = "��ͣ������"
  Timer1.Enabled = True
  OnTest = True
  '�������ģ��״̬
  Else
  Command1.Caption = "��һ��"
  Timer1.Enabled = False
  OnTest = False
  End If
  End Sub
    
Private Sub Command2_Click()
Print "sb"
End Sub

  '�������ʱһЩ������Ҫ����
  Private Sub Form_Load()
  Option_Tag = 1
  Timer1.Enabled = False
  OnTest = False
  End Sub
    
 
    
  'ÿ��һ����ģ��һ������¼�
  Private Sub Timer1_Timer()
  If Option_Tag = 1 Then
          '������mouse_event����������������ü�ǰ��˵��
          '���ͬʱҪģ����������¼���������   Or   ��������������
          '������   ����������   ���ɿ������¼�����ϼ�һ�ε���
          mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  ElseIf Option_Tag = 2 Then
          'ģ������Ҽ������¼�
          mouse_event MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
  Else
          '���������������������¼�   ����һ�����˫���¼�
          mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
          mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  End If
  End Sub

