VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "��Դ����ļ���ȡ"
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
'impactX Game Engine
'written by Davy.xu
'��res.grf����Դ�ļ����ж�ȡ�ļ�
'���ڲ˵� ����->��������� DirectX 8 for Visual Basic Type Library
Option Explicit
Dim pic As New xGraphPool

Private Sub Form_Load()
    Me.Width = 800 * 15 '���ڴ��ڴ�С
    Me.Height = 600 * 15
    InitDXGraph 800, 600, Me.hWnd, xgWindow '��ʼ��DXGraph
    '��ȡbmp
    pic.LoadGraphFromRes "res.grf", "bomb.bmp", xgBlack
    Me.Show
    Do
        DoEvents '��Windows���������
        PaintScreen 0 '�Ժ�ɫ������Ļ
        RenderBegin '��ʼ��Ⱦ
        pic.DrawGraph 100, 100
        RenderEnd '������Ⱦ
    Loop
End Sub
'����������ʱж��DX
Private Sub Form_Unload(Cancel As Integer)
    pic.Release
    UnloadDXGraph
    End
End Sub

