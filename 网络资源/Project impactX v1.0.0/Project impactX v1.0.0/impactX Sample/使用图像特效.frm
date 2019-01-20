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
'一些图像特效的演示
'请在菜单 工程->引用中添加 DirectX 8 for Visual Basic Type Library
Option Explicit
Dim mc As New xGraphPool
Dim wlogo As New xGraphPool
Dim bomb As New xGraphPool
Private Sub Form_Load()
    Me.Width = 800 * 15 '调节窗口大小
    Me.Height = 600 * 15
    InitDXGraph 800, 600, Me.hWnd, xgWindow '初始化DXGraph
    
    mc.LoadGraph "mcircle.bmp", xgBLACK
    wlogo.LoadGraph "warlogo.jpg", xgBLACK
    bomb.LoadGraph "bomb.bmp", xgBLACK, 6, 3 '设置图片的分割参数:横向6张，纵向3张
    
    Me.Show
    Do
        DoEvents '让Windows做别的事情
        PaintScreen 0 '以黑色擦除屏幕
        RenderBegin '开始渲染
        '不同的绘制顺序将会导致图片重叠效果
        Call ColorBlend
        Call Rotate
        Call CellControl
        
        LimitFPS 80 '限制FPS
        RenderEnd '结束渲染
    Loop
End Sub
'但窗口销毁时卸载DX
Private Sub Form_Unload(Cancel As Integer)
    UnloadDXGraph
    mc.Release '释放空间
    wlogo.Release
    bomb.Release
    End
End Sub
'旋转
Public Sub Rotate()
    Static Angle As Integer
    Angle = Angle + 1
    If Angle > 360 Then Angle = 0
    mc.SetRotate Angle
    mc.SetAlpha 128
    mc.DrawGraph 100, 100
End Sub
'颜色渲染
Public Sub ColorBlend()
    wlogo.SetColor D3DColorARGB(255, 100, 100, 255)
    wlogo.DrawGraph 120, 120
End Sub
'帧单元控制
Public Sub CellControl()
    Static Frame As Single
    Frame = Frame + 0.05 '播放速度
    If Frame > 18 Then Frame = 0
    bomb.Cell = Int(Frame)
    bomb.DrawGraph 100, 100
    DrawText "播放帧:" & bomb.Cell, 100, 80
End Sub
