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
'һЩͼ����Ч����ʾ
'���ڲ˵� ����->��������� DirectX 8 for Visual Basic Type Library
Option Explicit
Dim mc As New xGraphPool
Dim wlogo As New xGraphPool
Dim bomb As New xGraphPool
Private Sub Form_Load()
    Me.Width = 800 * 15 '���ڴ��ڴ�С
    Me.Height = 600 * 15
    InitDXGraph 800, 600, Me.hWnd, xgWindow '��ʼ��DXGraph
    
    mc.LoadGraph "mcircle.bmp", xgBLACK
    wlogo.LoadGraph "warlogo.jpg", xgBLACK
    bomb.LoadGraph "bomb.bmp", xgBLACK, 6, 3 '����ͼƬ�ķָ����:����6�ţ�����3��
    
    Me.Show
    Do
        DoEvents '��Windows���������
        PaintScreen 0 '�Ժ�ɫ������Ļ
        RenderBegin '��ʼ��Ⱦ
        '��ͬ�Ļ���˳�򽫻ᵼ��ͼƬ�ص�Ч��
        Call ColorBlend
        Call Rotate
        Call CellControl
        
        LimitFPS 80 '����FPS
        RenderEnd '������Ⱦ
    Loop
End Sub
'����������ʱж��DX
Private Sub Form_Unload(Cancel As Integer)
    UnloadDXGraph
    mc.Release '�ͷſռ�
    wlogo.Release
    bomb.Release
    End
End Sub
'��ת
Public Sub Rotate()
    Static Angle As Integer
    Angle = Angle + 1
    If Angle > 360 Then Angle = 0
    mc.SetRotate Angle
    mc.SetAlpha 128
    mc.DrawGraph 100, 100
End Sub
'��ɫ��Ⱦ
Public Sub ColorBlend()
    wlogo.SetColor D3DColorARGB(255, 100, 100, 255)
    wlogo.DrawGraph 120, 120
End Sub
'֡��Ԫ����
Public Sub CellControl()
    Static Frame As Single
    Frame = Frame + 0.05 '�����ٶ�
    If Frame > 18 Then Frame = 0
    bomb.Cell = Int(Frame)
    bomb.DrawGraph 100, 100
    DrawText "����֡:" & bomb.Cell, 100, 80
End Sub
