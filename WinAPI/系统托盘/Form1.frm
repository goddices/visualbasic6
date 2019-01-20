VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu mnuTray 
      Caption         =   "SS"
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuTrayClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuTrayMaximize 
         Caption         =   "Maximize"
      End
      Begin VB.Menu mnuTrayMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuTrayMove 
         Caption         =   "Move"
      End
      Begin VB.Menu mnuTrayRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuTraySize 
         Caption         =   "Size"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------
'           ʹ��ϵͳ���̳�����ʾ
'---------------------------------------------
'           ������� ��֪����
'---------------------------------------------
'����˵����
'   ����һ���Ƚ�������ʹ��ϵͳ���̵ĳ���ʵ��������
'�ˣ��������ͼ�꣬ɾ������ͼ�꣬��̬�ı�����ͼ�꣬
'Ϊ����ͼ����Ӹ�����ʾ��Ϣ��ʵ������ͼ�������Ҽ�
'�˵������ݡ�
'-------����-------------------����------------
'       Form1                   ������
'       mnuFile,mnuFileExit     �ļ��˵����˵���
'       mnuTray,mnuTrayClose... �������Ҽ��˵����˵���
'---------------------------------------------

Option Explicit

'LastState�����������Ǳ�ʾ������ԭ��״̬
Public LastState As Integer

'��VB������
'  Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'��˵����
'  ����һ�����ڵĴ��ں�������һ����Ϣ�����Ǹ����ڡ�������Ϣ������ϣ�����ú������᷵�ء�SendMessageBynum��
'  SendMessageByString�Ǹú����ġ����Ͱ�ȫ��������ʽ

'������ֵ��
'  Long���ɾ������Ϣ����

'��������
'  hwnd -----------  Long��Ҫ������Ϣ���Ǹ����ڵľ��

'  wMsg -----------  Long����Ϣ�ı�ʶ��

'  wParam ---------  Long������ȡ������Ϣ

'  lParam ---------  Any������ȡ������Ϣ
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'��ʾ���͵���ϵͳ����
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&

'�����������ʱ
Private Sub Form_Load()
    
    '�����WindowState���ԣ����ػ�����һ��ֵ����ֵ����ָ��������ʱ���崰�ڵĿ���״̬
    'vbNormal    0   ��ȱʡֵ������ ��
    'VbMinimized 1   ��С������С��Ϊһ��ͼ�꣩
    'VbMaximized 2   ��󻯣��������ߴ磩
    If WindowState = vbMinimized Then
        LastState = vbNormal
    Else
        LastState = WindowState
    End If
    
    '��ͼ����ӵ����̵ĺ������μ�ģ���еĽ���
    'ע�������Ǵ�������ģ�����ڣ������в�û��ֱ�ӵ���Shell_NotifyIcon����
    AddToTray Me, mnuTray
    
    SetTrayTip "����ͼ����ʾ������Ҽ������˵�"
End Sub

'��������Form1��С�ı�ʱ����Ӧ�ı��Ҽ��˵�mnuTray�Ĳ˵���Ŀ�������Enabled
Private Sub Form_Resize()
    Select Case WindowState
        
        '���������С���ˣ��Ѳ˵����󻯡����ָ�����Ϊ���ã�
        '���ѡ���С�������ƶ�������С��������Ϊ������.
        '�����ʱ������ͼ���ϵ������Ҽ����ᷢ�ֲ��������Ϊ��ɫ
        Case vbMinimized
            mnuTrayMaximize.Enabled = True
            mnuTrayMinimize.Enabled = False
            mnuTrayMove.Enabled = False
            mnuTrayRestore.Enabled = True
            mnuTraySize.Enabled = False
        
        '�������ʱ
        Case vbMaximized
            mnuTrayMaximize.Enabled = False
            mnuTrayMinimize.Enabled = True
            mnuTrayMove.Enabled = False
            mnuTrayRestore.Enabled = True
            mnuTraySize.Enabled = False
        
        'һ��״̬��
        Case vbNormal
            mnuTrayMaximize.Enabled = True
            mnuTrayMinimize.Enabled = True
            mnuTrayMove.Enabled = True
            mnuTrayRestore.Enabled = False
            mnuTraySize.Enabled = True
    End Select

    If WindowState <> vbMinimized Then LastState = WindowState
End Sub

'��֤�ڳ����˳�ʱɾ������ͼ��
Private Sub Form_Unload(Cancel As Integer)
    RemoveFromTray
End Sub

'���ļ����˵��ġ��˳�������ʱ
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

'����ͼ���Ҽ��˵��ϵġ��˳�������ʱ
Private Sub mnuTrayClose_Click()
    Unload Me
End Sub

'����ͼ���Ҽ��˵��ϵġ���󻯡�����ʱ
Private Sub mnuTrayMaximize_Click()
    WindowState = vbMaximized
End Sub

'����ͼ���Ҽ��˵��ϵġ���С��������ʱ
Private Sub mnuTrayMinimize_Click()
    WindowState = vbMinimized
End Sub

'����ͼ���Ҽ��˵��ϵġ��ƶ�������ʱ
Private Sub mnuTrayMove_Click()
    SendMessage HWnd, WM_SYSCOMMAND, _
        SC_MOVE, 0&
End Sub

'����ͼ���Ҽ��˵��ϵġ��ָ�������ʱ
Private Sub mnuTrayRestore_Click()
    SendMessage HWnd, WM_SYSCOMMAND, _
        SC_RESTORE, 0&
End Sub

'����ͼ���Ҽ��˵��ϵġ��˳�������ʱ
Private Sub mnuTraySize_Click()
    SendMessage HWnd, WM_SYSCOMMAND, _
        SC_SIZE, 0&
End Sub

 
