VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "���ƶ����ͬ����"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pic As New xGraphPool

Private Sub Form_Load()
    Me.Width = 800 * 15 '���ڴ��ڴ�С
    Me.Height = 600 * 15
    InitDXGraph 800, 600, Me.hWnd, xgWindow '��ʼ��DXGraph
    '��ȡbmp
    pic.LoadGraph "a.png", xgBLACK
    
    Me.Show
    Dim i As Integer
    Dim j As Integer
    Do
        DoEvents '��Windows���������
        PaintScreen 0 '�Ժ�ɫ������Ļ
        RenderBegin '��ʼ��Ⱦ
        'impactX��ÿһ��ͼ����Դ��һ��xGraphPoolʵ�������Զ���RPG�е����ֵ�
        '��ͬͼƬ�ĳ������ƣ�û�б�ҪΪÿһ����������һ��xGraphPool
        'ֻ��Ҫ����Щ�������걣�����һ��xGraphPool���Ƽ���
        For i = 0 To 300 Step 70
            For j = 0 To 300 Step 70
                pic.DrawGraph i, j
            Next j
        Next i
        
        RenderEnd '������Ⱦ
    Loop
End Sub
'����������ʱж��DX
Private Sub Form_Unload(Cancel As Integer)
    pic.Release
    UnloadDXGraph
    End
End Sub

