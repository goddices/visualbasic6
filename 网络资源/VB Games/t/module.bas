Attribute VB_Name = "Module1"
'*************************************************
'
'��ת����˹ 1.0 Demo Դ����
'̩����������� ��ǿ �� 1998��4�� �ṩ
'������κβ���֮�����������
'http://www.nease.net/~jackyyin �����۰��������
'
'*************************************************
Global CurX As Integer            'ĿǰX����
Global Total(10, 20) As Boolean    '�������겼�� 10x20
Global MinX As Integer 'һ���������� x ����
Global MaxX As Integer 'һ���������С x ����
Global MinY As Integer 'һ���������� y ����
Global MaxY As Integer 'һ���������С y ����

Type cXs    'һ������ 4 ���������
    cX As Integer 'x ����
    cY As Integer 'y ����
    cZ As Boolean '�ж�һ���������Ƿ��ǿյ�
End Type
Global Xs(4) As cXs

Global Adjust_Left As Integer '��ת�����󷽵�����λ��
Global Adjust_Top As Integer  '��ת�����Ϸ�������λ��
    
'BitBlt �������ã�λ����λͼ��ʵ�ֲ�����ķ���Ķ���
Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

