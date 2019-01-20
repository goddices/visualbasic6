VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "第一个DirectGraph程序"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'impactX Game Engine
'written by Davy.xu
'创建一个没有图片的DX程序
'请在菜单 工程->引用中添加 DirectX 8 for Visual Basic Type Library
Option Explicit

Private Sub Form_Load()
    Me.Width = 800 * 15 '调节窗口大小
    Me.Height = 600 * 15
    InitDXGraph 800, 600, Me.hWnd, xgWindow '初始化DXGraph
    Me.Show
    Do
        DoEvents '让Windows做别的事情
        PaintScreen 0 '以黑色擦除屏幕
        RenderBegin '开始渲染
            DrawCircle 200, 200, 200, xgWHITE '画一个圆
            DrawRectFill 100, 100, 300, 300, D3DColorARGB(128, 255, 255, 0)
            DrawRectFill 200, 200, 400, 400, D3DColorARGB(128, 0, 0, 255)
        RenderEnd '结束渲染
    Loop
End Sub
'但窗口销毁时卸载DX
Private Sub Form_Unload(Cancel As Integer)
    UnloadDXGraph
    End
End Sub
