VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��̬ͼƬ�ؼ�"
   ClientHeight    =   6480
   ClientLeft      =   2340
   ClientTop       =   2940
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   8340
   Begin VB.CommandButton Command2 
      Caption         =   "��     ��"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��     ��"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   11
      Interval        =   10
      Left            =   7800
      Top             =   120
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   10
      Interval        =   10
      Left            =   7320
      Top             =   120
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   9
      Interval        =   10
      Left            =   6840
      Top             =   120
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   10
      Left            =   6360
      Top             =   120
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   7
      Interval        =   10
      Left            =   5880
      Top             =   120
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   10
      Left            =   5400
      Top             =   120
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   10
      Left            =   4920
      Top             =   120
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   10
      Left            =   4440
      Top             =   120
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   10
      Left            =   3960
      Top             =   120
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   10
      Left            =   3480
      Top             =   120
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   10
      Left            =   3000
      Top             =   120
   End
   Begin VB.PictureBox P1 
      Height          =   3660
      Left            =   960
      ScaleHeight     =   3600
      ScaleMode       =   0  'User
      ScaleWidth      =   2415
      TabIndex        =   1
      Top             =   960
      Width           =   2475
   End
   Begin VB.PictureBox P2 
      AutoSize        =   -1  'True
      Height          =   3660
      Left            =   4800
      ScaleHeight     =   3600
      ScaleMode       =   0  'User
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   960
      Width           =   2475
   End
   Begin VB.Label Label2 
      Caption         =   "������ߣ�¬����"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "��Ȩ���У�����LPP���������"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Menu trick 
      Caption         =   "�Ƴ�"
      Index           =   1
   End
   Begin VB.Menu trick 
      Caption         =   "����"
      Index           =   2
   End
   Begin VB.Menu trick 
      Caption         =   "��Ҷ��"
      Index           =   3
      Begin VB.Menu window 
         Caption         =   "��ֱˮƽ��Ҷ��"
         Index           =   1
      End
      Begin VB.Menu window 
         Caption         =   "�����Ҷ��"
         Index           =   2
      End
      Begin VB.Menu window 
         Caption         =   "�����Ҷ��"
         Index           =   3
      End
   End
   Begin VB.Menu trick 
      Caption         =   "������"
      Index           =   4
   End
   Begin VB.Menu trick 
      Caption         =   "��Ļ"
      Index           =   5
      Begin VB.Menu curtain 
         Caption         =   "��������������Ļ"
         Index           =   1
      End
      Begin VB.Menu curtain 
         Caption         =   "��������������Ļ"
         Index           =   2
      End
      Begin VB.Menu curtain 
         Caption         =   "��������������Ļ"
         Index           =   3
      End
   End
   Begin VB.Menu trick 
      Caption         =   "����"
      Index           =   6
      Begin VB.Menu roll 
         Caption         =   "���ұ�����߹���"
         Index           =   1
      End
      Begin VB.Menu roll 
         Caption         =   "�ϰ벿���°벿��λ����"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  

   
    
  Public w, h, x, y
    
  Private Sub Command1_Click()
  For i = 1 To 11
      Timer(i).Enabled = False
  Next i
  P1.Cls
  x = 0
  y = 0
  End Sub
    
  Private Sub Command2_Click()
  End
  End Sub
    
  Private Sub curtain_Click(Index As Integer)
  Select Case Index
      Case 1
          x = w / 2
          Timer(7).Enabled = True
      Case 2
          x = w / 2
          y = h / 2
          Timer(8).Enabled = True
      Case 3
          Timer(9).Enabled = True
  End Select
  End Sub
    
  Private Sub Form_Load()
  w = P2.Width
  h = P2.Height
  End Sub
    
  Private Sub roll_Click(Index As Integer)
  Select Case Index
      Case 1
          Timer(10).Enabled = True
      Case 2
          Timer(11).Enabled = True
  End Select
  End Sub
    
  Private Sub Timer_Timer(Index As Integer)
  If Timer(Index).Enabled = False Then
      x = 0
      y = 0
  End If
  Select Case Index
      Case 1                                   '�Ƴ�       (x,y��ʼֵΪͼƬ���һ��)
          x = x - w / 100
          y = y - h / 100
          If x <= 0 Or y <= 0 Then Timer(1).Enabled = False
          P1.PaintPicture P2.Picture, 0, 0, w, h, x, y, w - 2 * x, h - 2 * y
      Case 2                                   '����
          P1.PaintPicture P2.Picture, 0, 0, w, h, x, y, w - 2 * x, h - 2 * y
          x = x + w / 100
          y = y + h / 100
          If x >= w / 3 Or y >= h / 3 Then Timer(2).Enabled = False
      Case 3                                   '��ֱˮƽ��Ҷ��
          m = w / 20:           n = h / 20
          x = x + w / 100:               y = y + h / 100
          If x >= m + 5 Or y >= n + 5 Then Timer(3).Enabled = False
          For i = 0 To 20
              For j = 0 To 20
                  P1.PaintPicture P2.Picture, i * m, j * n, , , i * m, j * n, x, y
              Next j
          Next i
      Case 4                                   '�����Ҷ��
          n = h / 20
          y = y + h / 100
          If y >= h Then Timer(4).Enabled = False
          For i = 0 To 20
              P1.PaintPicture P2.Picture, 0, i * n, , , 0, i * n, w, y
          Next i
      Case 5                                   '�����Ҷ��
          m = w / 20
          x = x + w / 100
          If x >= w Then Timer(5).Enabled = False
          For j = 0 To 20
              P1.PaintPicture P2.Picture, j * m, 0, , , j * m, 0, x, h
          Next j
      Case 6                                   '������
          c = c + 1
          If c > 100 Then P1.PaintPicture P2.Picture, 0, 0:                   c = 0
          m = w / 50:           n = h / 50
          For i = 1 To 50 + c * 10
              xx = Rnd * (w - m - 50)
              yy = Rnd * (h - n - 50)
              P1.PaintPicture P2.Picture, xx, yy, , , xx, yy, m, n
          Next i
      Case 7                                   '��������������Ļ     (x��ʼֵΪͼƬ���һ��)
          x = x - 10
          If x <= 0 Then Timer(7).Enabled = False
          P1.PaintPicture P2.Picture, x, 0, w - 2 * x, h, x, 0, w - 2 * x, h
      Case 8                                   '��������������Ļ       (x,y��ʼֵΪͼƬ���һ��)
          x = x - w / 100
          y = y - h / 100
          If x <= 0 Or y <= 0 Then Timer(8).Enabled = False
          P1.PaintPicture P2.Picture, x, y, w - 2 * x, h - 2 * y, x, y, w - 2 * x, h - 2 * y
      Case 9                                   '��������������Ļ
          x = x + w / 100
          y = y + h / 100
          If x <= 0 Or y <= 0 Then Timer(9).Enabled = False
          P1.PaintPicture P2.Picture, 0, 0, w, y, 0, 0, w, y
          P1.PaintPicture P2.Picture, 0, 0, x, h, 0, 0, x, h
          P1.PaintPicture P2.Picture, 0, h - y, w, h - y, 0, h - y, w, h - y
          P1.PaintPicture P2.Picture, w - x, 0, w - x, h, w - x, 0, w - x, h
      Case 10                                 '���ұ�����߹���
          x = x + w / 100
          If x >= w Then Timer(10).Enabled = False
          P1.PaintPicture P2.Picture, w - x, 0, x, h, 0, 0, x, h
      Case 11                                 '�ϰ벿���°벿��λ����
          x = x + w / 100
          If x >= w Then Timer(11).Enabled = False
          P1.PaintPicture P2.Picture, w - x, 0, , , 0, 0, x, h / 2
          P1.PaintPicture P2.Picture, 0, h / 2, , , w - x, h / 2, x, h / 2
  End Select
    
  End Sub
    
  Private Sub trick_Click(Index As Integer)
  Select Case Index
      Case 1
          x = w / 2:           y = h / 2
          Timer(1).Enabled = True
      Case 2
          Timer(2).Enabled = True
      Case 4
          Timer(6).Enabled = True
  End Select
  End Sub
    
  Private Sub window_Click(Index As Integer)
  Select Case Index
      Case 1
          Timer(3).Enabled = True
      Case 2
          Timer(4).Enabled = True
      Case 3
          Timer(5).Enabled = True
  End Select
  End Sub
  
