VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "���ͼ��̵�ʹ��"
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
'���ڲ˵� ����->��������� '���ڲ˵� ����->��������� DirectX 8 for Visual Basic Type Library
Option Explicit
Dim xi As New xInput
Private Sub Form_Load()
    Me.Width = 800 * 15 '���ڴ��ڴ�С
    Me.Height = 600 * 15
    InitDXGraph 800, 600, Me.hWnd, xgWindow '��ʼ��DXGraph
    xi.InitDXInput Me.hWnd '��ʼ��DXinput
    Me.Show
    Do
        DoEvents '��Windows���������
        PaintScreen 0 '�Ժ�ɫ������Ļ
        RenderBegin '��ʼ��Ⱦ
            DrawText "�������:" & xi.MouseX & "," & xi.MouseY, 30, 10, xgWHITE
            DrawText IIf(xi.KeyInput(DIK_RETURN), "���»س���", "�밴�س���"), 30, 30, xgGREEN
            DrawText IIf(xi.MouseKey(xgL_BUTTON), "����������", "�밴������"), 30, 50, xgGREEN
            
        RenderEnd '������Ⱦ
    Loop
End Sub
'����������ʱж��DX
Private Sub Form_Unload(Cancel As Integer)
    UnloadDXGraph
    End
End Sub
