VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "��һ��DirectGraph����"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '��Ļ����
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'impactX Game Engine
'written by Davy.xu
'����һ��û��ͼƬ��DX����
'���ڲ˵� ����->��������� DirectX 8 for Visual Basic Type Library
Option Explicit

Private Sub Form_Load()
    Me.Width = 800 * 15 '���ڴ��ڴ�С
    Me.Height = 600 * 15
    InitDXGraph 800, 600, Me.hWnd, xgWindow '��ʼ��DXGraph
    Me.Show
    Do
        DoEvents '��Windows���������
        PaintScreen 0 '�Ժ�ɫ������Ļ
        RenderBegin '��ʼ��Ⱦ
            DrawCircle 200, 200, 200, xgWHITE '��һ��Բ
            DrawRectFill 100, 100, 300, 300, D3DColorARGB(128, 255, 255, 0)
            DrawRectFill 200, 200, 400, 400, D3DColorARGB(128, 0, 0, 255)
        RenderEnd '������Ⱦ
    Loop
End Sub
'����������ʱж��DX
Private Sub Form_Unload(Cancel As Integer)
    UnloadDXGraph
    End
End Sub
