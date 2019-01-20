Attribute VB_Name = "Module1"
'*************************************************
'
'旋转俄罗斯 1.0 Demo 源程序
'泰立软件工作室 尹强 于 1998年4月 提供
'如果有任何不明之处，你可以上
'http://www.nease.net/~jackyyin 的讨论版进行讨论
'
'*************************************************
Global CurX As Integer            '目前X坐标
Global Total(10, 20) As Boolean    '总体坐标布局 10x20
Global MinX As Integer '一个方块的最大 x 坐标
Global MaxX As Integer '一个方块的最小 x 坐标
Global MinY As Integer '一个方块的最大 y 坐标
Global MaxY As Integer '一个方块的最小 y 坐标

Type cXs    '一个方块 4 个点的坐标
    cX As Integer 'x 坐标
    cY As Integer 'y 坐标
    cZ As Boolean '判断一个点下面是否是空的
End Type
Global Xs(4) As cXs

Global Adjust_Left As Integer '翻转后向左方调整的位置
Global Adjust_Top As Integer  '翻转后向上方调整的位置
    
'BitBlt 函数作用：位操作位图，实现不规则的方块的动作
Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

