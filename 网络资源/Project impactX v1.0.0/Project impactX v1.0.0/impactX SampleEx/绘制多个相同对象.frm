VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "绘制多个相同对象"
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
    Me.Width = 800 * 15 '调节窗口大小
    Me.Height = 600 * 15
    InitDXGraph 800, 600, Me.hWnd, xgWindow '初始化DXGraph
    '读取bmp
    pic.LoadGraph "a.png", xgBLACK
    
    Me.Show
    Dim i As Integer
    Dim j As Integer
    Do
        DoEvents '让Windows做别的事情
        PaintScreen 0 '以黑色擦除屏幕
        RenderBegin '开始渲染
        'impactX的每一个图形资源是一个xGraphPool实例，所以对于RPG中的树林等
        '相同图片的场景绘制，没有必要为每一棵树都申请一个xGraphPool
        '只需要将这些树的坐标保存后用一个xGraphPool绘制即可
        For i = 0 To 300 Step 70
            For j = 0 To 300 Step 70
                pic.DrawGraph i, j
            Next j
        Next i
        
        RenderEnd '结束渲染
    Loop
End Sub
'但窗口销毁时卸载DX
Private Sub Form_Unload(Cancel As Integer)
    pic.Release
    UnloadDXGraph
    End
End Sub

