VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "资源打包文件读取"
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
'从res.grf的资源文件包中读取文件
'请在菜单 工程->引用中添加 DirectX 8 for Visual Basic Type Library
Option Explicit
Dim pic As New xGraphPool

Private Sub Form_Load()
    Me.Width = 800 * 15 '调节窗口大小
    Me.Height = 600 * 15
    InitDXGraph 800, 600, Me.hWnd, xgWindow '初始化DXGraph
    '读取bmp
    pic.LoadGraphFromRes "res.grf", "bomb.bmp", xgBlack
    Me.Show
    Do
        DoEvents '让Windows做别的事情
        PaintScreen 0 '以黑色擦除屏幕
        RenderBegin '开始渲染
        pic.DrawGraph 100, 100
        RenderEnd '结束渲染
    Loop
End Sub
'但窗口销毁时卸载DX
Private Sub Form_Unload(Cancel As Integer)
    pic.Release
    UnloadDXGraph
    End
End Sub

