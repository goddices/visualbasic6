Attribute VB_Name = "Mod_gol"

'Physics Toy V2.0

'Physics Toy(物理玩具)是本人设计的一款2D物理模拟软件,仿制了真实环境物体的各种连接与碰撞效果,有需要的可与我联系

'应用介绍：
'少年智力开发，学校物理课程指导，各种机构运动的研究

'基于冲量的刚体铰链与碰撞,销限制铰接计算出链、摆、小车等模型,通过与扭簧配合模拟出玩具木马模型
'仿真真实世界物理状态，建立了可用户控制的功能,达成了如同儿童玩具室的环境

'                                                         www.vbgamedev.com
'                                                               2007.5.5  by zh1110  China

Option Explicit

'对象定义，初始化模块
Public Const Def_Elasticity As Single = 0.4   '默认的反弹系数
Public Const Def_Friction As Single = 0.6    '默认的摩擦系数

Public CUR_Friction As Single

Public Enum TAPE ' 定义刚体类型
    TAPEBOX = 1
    TAPECIRCLE = 2
End Enum

Public Type Rig_Body
    TAPE As Long
    m As Single '质量
    pos As D3DVECTOR '位置
    v As D3DVECTOR '速度
    ang  As Single '角度
    Vang As D3DVECTOR '角速度
    Iz As Single '惯性矩

    Rbou As Single '包围球半径(圆类型即半径)
    Friction As Single '摩擦系数
    BColor As Long '颜色

    '以下刚体为圆类型时不用
    w As Single '宽度
    h As Single '高度
    tvertes()   As D3DVECTOR '4角点存储
    vertes()   As D3DVECTOR '4角点
    nor() As D3DVECTOR
    d() As Single
    numverts As Long
    ID As Long
End Type

Public Type collision
    Pn As D3DVECTOR '冲量
    Pt As Single  '摩擦冲量
    tangent_vel As D3DVECTOR '摩擦速度
    tangent_speed As Single  '摩擦速度大小

    Vn As Single '相对速度
    Vn2 As Single '撞击后的相对速度
    N As D3DVECTOR '碰撞垂直方向

    Ra  As D3DVECTOR '撞击点到重心向量(刚体a)
    Va As D3DVECTOR '撞击点的速度(刚体a)
    Ca As Single  '常量因子(刚体a)

    Rb  As D3DVECTOR '撞击点到重心向量(刚体b)
    Vb As D3DVECTOR '撞击后撞击点的速度(刚体B)
    Cb As Single  '常量因子(刚体B)

    c As Single  '常量因子(Ca+Cb)
End Type

Public Type Cline
    starpnt As D3DVECTOR
    endpnt As D3DVECTOR
    nor    As D3DVECTOR
    d As Single
End Type

'铰接
Public Type Constraint_body_body_point
    body0_pos   As D3DVECTOR
    body1_pos   As D3DVECTOR
    NUMbody0 As Long
    NUMbody1 As Long
    Friction_Attenuation As Single '铰接处的衰减摩擦转矩
End Type

'直槽
Public Type Constraint_SLOT
    body0_0pos   As D3DVECTOR
    body0_1pos   As D3DVECTOR
    body1_pos   As D3DVECTOR
    NUMbody0 As Long
    NUMbody1 As Long
    Restrictions As Boolean '是否限制在两端点内
End Type

'线性弹簧
Public Type LinearSpring
    body0_pos   As D3DVECTOR
    body1_pos   As D3DVECTOR
    NUMbody0 As Long
    NUMbody1 As Long
    k As Single '刚度
    FreeLong As Single '自由长
End Type

'扭簧
Public Type TwistSpring
    BaseAng As Single '基角
    k As Single '刚度
    NUMbody0 As Long
    NUMbody1 As Long
End Type

'转动电马达
Public Type TwistMotor
    TOR As Single 'O速度转矩
    max_angv As Single  '最大角速度
    NUMbody0 As Long
    NUMbody1 As Long
End Type

Public NUMBox As Long
Public box() As Rig_Body

'铰接
Public NUMJoint As Long
Public Joint() As Constraint_body_body_point

'直槽
Public NUMSLOT As Long
Public Slots() As Constraint_SLOT

'线性弹簧
Public NUMLinSpring As Long
Public LinSprings() As LinearSpring

'扭簧
Public NUMTwSpring As Long
Public TwSpring() As TwistSpring

'转动电马达
Public NUMTwMotor As Long
Public TwMotor() As TwistMotor

'墙
Public NUM_WALL As Long
Public WALL() As Cline

Public COLMAP() As Boolean  '刚体的碰撞检测查找表

Public force_CE As Boolean
Public Mouse_x As Single, Mouse_y As Single
Public MouseHit_Num As Long
Public MouseHit_pos As D3DVECTOR
Public Mouse_pos As D3DVECTOR
Public MouseJoint As Constraint_body_body_point

Public Sub Main()
    Randomize
    Form1.Show
    FormDEBUG.Show

 Call Scene1
    Call mainloop
End Sub

Public Sub Rclear()
    NUMBox = 0
    NUMJoint = 0
    NUMSLOT = 0
    NUMLinSpring = 0
    NUMTwSpring = 0
    NUMTwMotor = 0
    NUM_WALL = 0
    Call Rrsort
End Sub

Public Sub Rrsort()
    ReDim box(NUMBox)
    ReDim COLMAP(NUMBox, NUMBox)
    ReDim Joint(NUMJoint)
    ReDim Slots(NUMSLOT)
    ReDim LinSprings(NUMLinSpring)

    ReDim TwSpring(NUMTwSpring)
    ReDim TwMotor(NUMTwMotor)

    ReDim WALL(NUM_WALL)
End Sub

'创建正方形刚体
Public Sub CREATBOX(NUM As Long, BodW As Single, BodH As Single)
    box(NUM).ID = NUM
    box(NUM).TAPE = TAPEBOX
    box(NUM).w = BodW
    box(NUM).h = BodH

    If box(NUM).m = 0 Then box(NUM).m = 1 * BodW * BodH
    If box(NUM).Iz = 0 Then box(NUM).Iz = box(NUM).m * (BodW ^ 2 + BodH ^ 2) / 12
    If box(NUM).Rbou = 0 Then box(NUM).Rbou = Sqr((BodW / 2) ^ 2 + (BodH / 2) ^ 2)
    If box(NUM).Friction = 0 Then box(NUM).Friction = Def_Friction
    '    If (box(NUM).v.X = 0 And box(NUM).v.Y = 0) Then box(NUM).v = Makever(0, 0, 0)
    '    If box(NUM).ang = 0 Then box(NUM).ang = 0
    '    If box(NUM).Vang.z = 0 Then box(NUM).Vang.z = 0
    box(NUM).numverts = 4
    ReDim box(NUM).tvertes(1 To 4)
    ReDim box(NUM).vertes(1 To 4)
    ReDim box(NUM).nor(1 To 4)
    ReDim box(NUM).d(1 To 4)
    '4 _______ 1
    ' |       |
    ' |       |
    '3|_______|2

    box(NUM).tvertes(1) = Makever(box(NUM).w / 2, box(NUM).h / 2, 0)
    box(NUM).tvertes(2) = Makever(box(NUM).w / 2, -box(NUM).h / 2, 0)
    box(NUM).tvertes(3) = Makever(-box(NUM).w / 2, -box(NUM).h / 2, 0)
    box(NUM).tvertes(4) = Makever(-box(NUM).w / 2, box(NUM).h / 2, 0)
End Sub

'创建三角形刚体
Public Sub CREATTriangle(NUM As Long, A As Single, b As Single, c As Single)
    If (A >= b + c Or b >= A + c Or c >= A + b) Then
        show_debug "Triangle error" '两边之和大于第三边
        Exit Sub
    End If

    '三角形面积 海伦公式:S=sqrt[p(p-a)(p-b)(p-c)] , p=(a+b+c)/2
    Dim p As Single, S As Single
    p = ((A + b + c) / 2)
    S = Sqr(p * (p - A) * (p - b) * (p - c))

    box(NUM).ID = NUM
    box(NUM).TAPE = TAPEBOX
    box(NUM).m = 1 * S
    If box(NUM).Friction = 0 Then box(NUM).Friction = Def_Friction

    '       2
    '      /|\
    '     / | \
    '   b/  |   \c
    '   /   |h   \
    '  /    |      \
    '1/__a1_|_______\3
    '        a
    '余弦定理c2=a2+b2-2ab*cosQ

    Dim a1 As Single, h As Single, cosQ As Single
    cosQ = (A * A + b * b - c * c) / (2 * A * b)
    a1 = cosQ * b
    h = Sqr(b * b - a1 * a1)

    box(NUM).Iz = box(NUM).m * (A * A + a1 * a1 + h * h - A * a1) / 18

    box(NUM).numverts = 3
    ReDim box(NUM).tvertes(1 To 3)
    ReDim box(NUM).vertes(1 To 3)
    ReDim box(NUM).nor(1 To 3)
    ReDim box(NUM).d(1 To 3)

    Dim Focus_x As Single, Focus_y As Single '重心
    Focus_x = A / 2 - (A / 2 - a1) / 3
    Focus_y = h / 3
    box(NUM).tvertes(1) = Makever(-Focus_x, -Focus_y)
    box(NUM).tvertes(2) = Makever((a1 - A / 2) * 2 / 3, h * 2 / 3)
    box(NUM).tvertes(3) = Makever(A - Focus_x, -Focus_y)

    Dim ds(3) As Single
    ds(1) = VLength(box(NUM).tvertes(1))
    ds(2) = VLength(box(NUM).tvertes(2))
    ds(3) = VLength(box(NUM).tvertes(3))

    box(NUM).Rbou = MaxVel(MaxVel(ds(1), ds(2)), ds(2))
End Sub

'创建圆形刚体
Public Sub CREATCIRCLE(NUM As Long, BodR As Single)
    box(NUM).ID = NUM
    box(NUM).TAPE = TAPECIRCLE
    box(NUM).Rbou = BodR
    box(NUM).numverts = 0
    If box(NUM).Friction = 0 Then box(NUM).Friction = Def_Friction
    If box(NUM).m = 0 Then box(NUM).m = 1 * pi * box(NUM).Rbou ^ 2
    If box(NUM).Iz = 0 Then box(NUM).Iz = (box(NUM).m * box(NUM).Rbou ^ 2) / 2
End Sub

Public Sub CREATJoint(NUM As Long, NUMbody0 As Long, NUMbody1 As Long)
    Joint(NUM).NUMbody0 = NUMbody0
    Joint(NUM).NUMbody1 = NUMbody1
End Sub

Public Sub Make_Wall(NUM As Long, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single)
    With WALL(NUM)
        .starpnt = Makever(X1, Y1)
        .endpnt = Makever(X2, Y2)
        CaleLineN .starpnt, .endpnt, .nor, .d
    End With
End Sub

'玩具室
Public Sub Scene1()

    Rclear

    NUMBox = 19
    NUMJoint = 13
    NUMTwSpring = 6
    NUM_WALL = 4
    NUMLinSpring = 1

    Call Rrsort

    Dim k As Long, S As Long
    'Rigid_body 19车体  1前轮  2后轮
    Call CREATBOX(19, 5, 2.5)
    box(19).pos = Makever(7, 1)

    Call CREATCIRCLE(1, 1.2)
    box(1).pos = box(19).pos
    box(1).Friction = 0.9 '轮的摩擦稍大
    box(1).m = 0.5 * pi * box(1).Rbou ^ 2 '轮的质量稍小

    Call CREATCIRCLE(2, 1.2)
    box(2) = box(1) '轮的位置由铰接自动调整
    'box(2).BColor = RGB(0, 0, 200)

    'Rigid_body 3 ,4 ,5 ,6 ,7链
    For k = 3 To 7
        Call CREATBOX(k, 1, 3)
        box(k).pos = Makever(15 + k * 0.5, 22 - k * 2.5)
    Next

    'Rigid_body  8 吊摆盒
    Call CREATBOX(8, 2, 2.5)
    'Call CREATTriangle(8, 3, 4, 4.7)
    box(8).pos = Makever(14, 7)

    '    Rigid_body  9 物品盒子 10 11 两个小球
    '    Call CREATBOX(9, 2.2, 3.3)
    Call CREATTriangle(9, 4.2, 3.5, 5.5)
    box(9).pos = Makever(14.2, 12)

    '    Call CREATCIRCLE(10, 0.6)
    '    box(10).pos = Makever(14, 13)
    '
    '    Call CREATCIRCLE(11, 0.6)
    '    box(11).pos = Makever(14, 11)

    'Rigid_body 12大圆
    Call CREATCIRCLE(12, 1.5)
    box(12).pos = Makever(7, 13)
    box(12).BColor = &H606060

    'Rigid_body 13，14，15，16，17，18木马模型
    '躯干
    Call CREATBOX(13, 3, 2)
    box(13).pos = Makever(3, 15)

    '头
    Call CREATBOX(14, 1.6, 1.4)
    box(14).pos = box(13).pos

    '前腿
    Call CREATBOX(15, 2.2, 0.8)
    box(15).pos = box(13).pos
    box(15).Friction = 0.8 '腿的摩擦稍大

    '后腿
    Call CREATBOX(16, 2.2, 0.8)
    box(16) = box(15)

    '尾巴1
    Call CREATBOX(17, 1, 0.5)
    box(17).pos = box(13).pos

    '尾巴2
    Call CREATBOX(18, 1, 0.5)
    box(18) = box(17)

    For k = 13 To 18
        box(k).BColor = &H606060
        box(k).m = 0.6 * box(k).w * box(k).h
    Next

    COLMAP(13, 14) = 1

    '连接前轮
    With Joint(1)
        .NUMbody0 = 19
        .NUMbody1 = 1
        .body0_pos = Makever(2, -1)
        .body1_pos = Makever(0#, 0)
    End With
    COLMAP(1, 19) = 1

    '连接后轮
    With Joint(2)
        .NUMbody0 = 19
        .NUMbody1 = 2
        .body0_pos = Makever(-2, -1)
        .body1_pos = Makever(0#, 0)
    End With
    COLMAP(2, 19) = 1

    '连接铰链
    For k = 3 To 6
        With Joint(k)
            .NUMbody0 = k
            .NUMbody1 = k + 1
            .body0_pos = Makever(0.1, -1.5)
            .body1_pos = Makever(-0.1, 1.5)
            .Friction_Attenuation = 0.12
        End With
        COLMAP(k, k + 1) = 1 '取消相邻铰链的碰撞关系
    Next

    '连接摆
    With Joint(7)
        .NUMbody0 = 8
        .NUMbody1 = -1
        .body0_pos = Makever(0, 1.9)
        .body1_pos = Makever(14, 9)
        .Friction_Attenuation = 0.05
    End With

    show_debug "link Trojan"
    '连接木马模型各部分
    '连接头
    With Joint(8)
        .NUMbody0 = 13
        .NUMbody1 = 14
        .body0_pos = Makever(1.6, 1.2)
        .body1_pos = Makever(-0.6, -0.1)
    End With

    With TwSpring(1)
        .BaseAng = 0.4
        .k = 0.02
    End With

    '连接前腿
    With Joint(9)
        .NUMbody0 = 13
        .NUMbody1 = 15
        .body0_pos = Makever(1.3, -1.1)
        .body1_pos = Makever(-1, 0)
    End With

    With TwSpring(2)
        .BaseAng = -1.1
        .k = 0.37
    End With

    '连接后腿
    With Joint(10)
        .NUMbody0 = 13
        .NUMbody1 = 16
        .body0_pos = Makever(-1.3, -1.1)
        .body1_pos = Makever(1, 0)
    End With

    With TwSpring(3)
        .BaseAng = 1.1
        .k = 0.37
    End With

    '连接尾1
    With Joint(11)
        .NUMbody0 = 13
        .NUMbody1 = 17
        .body0_pos = Makever(-1.8, 0.8)
        .body1_pos = Makever(0.3, 0)
    End With

    With TwSpring(4)
        .BaseAng = 1.2
        .k = 0.005
    End With

    '连接尾2
    With Joint(12)
        .NUMbody0 = 17
        .NUMbody1 = 18
        .body0_pos = Makever(-0.5, 0)
        .body1_pos = Makever(0.5, 0)
    End With

    With TwSpring(5)
        .BaseAng = 0
        .k = 0.002
    End With

    For k = 1 To 5
        TwSpring(k).NUMbody0 = Joint(k + 7).NUMbody0
        TwSpring(k).NUMbody1 = Joint(k + 7).NUMbody1
    Next

    With LinSprings(1)
        .body0_pos = Makever(-1.2, 0)
        .body1_pos = Makever(1.4, 0.5)
        .NUMbody0 = 12
        .NUMbody1 = 13
        .FreeLong = 7
        .k = 0.02
    End With

    '取消模型各部分之间碰撞关系
    For k = 13 To 17
        For S = k + 1 To 18
            COLMAP(k, S) = 1
        Next
    Next

    '/////////////////////////////////////////////////
    Make_Wall 1, -2, -1, 10, -3
    Make_Wall 2, 10, -3, 21, -2.5
    Make_Wall 3, -1, 6, -2, -1
    Make_Wall 4, -1, 9, 8.5, 8.8

End Sub

'永动机
'在欧洲，早期最著名的一个永动机设计方案是十三世纪时一个叫亨内考的法国人提出来的。
'轮子中央有一个转动轴，轮子边缘安装着12个可活动的短杆，每个短杆的一端装有一个铁球。
'方案的设计者认为，右边的球比左边的球离轴远些，因此，右边的球产生的转动力矩要比左边的球产生的转动力矩大。
'这样轮子就会永无休止地沿着箭头所指的方向转动下去，并且带动机器转动。
'这个设计被不少人以不同的形式复制出来，但从未实现不停息的转动。
Public Sub Scene5()
    Dim k As Long, S As Long
    Rclear
    NUMSLOT = 12
    NUMJoint = 32
    NUMBox = 34
    NUM_WALL = 1
    Call Rrsort

    Make_Wall 1, -1, -3, 21, -3

    Call CREATCIRCLE(1, 5)
    box(1).pos = Makever(10, 6.5)

    With Joint(1)
        .NUMbody0 = 1
        .NUMbody1 = -1
        .body0_pos = Makever(0, 0)
        .body1_pos = box(1).pos
    End With

    Dim A As Single
    For k = 2 To 11
        A = PITWO * (k - 2) / 10

        box(k).m = 0.2
        Call CREATBOX(k, 5, 0.4) '10根主杆
        box(k).BColor = RGB(72, 72, 0)
        box(k).pos = Add(box(1).pos, Makever(5 * Cos(A), 5 * Sin(A)))
        box(k).ang = A

        With Joint(k)
            .NUMbody0 = 1
            .NUMbody1 = k
            .body0_pos = Makever(3 * Cos(A), 3 * Sin(A))
            .body1_pos = Makever(-2.4, 0)
            .Friction_Attenuation = 0.2
        End With

    Next

    For k = 12 To 21
        A = PITWO * (k - 12) / 10 - 0.1
        Call CREATCIRCLE(k, 0.4) '主盘上的10个销
        
        box(k).pos = Add(box(1).pos, Makever(4 * Cos(A), 4 * Sin(A)))
        With Joint(k)
            .NUMbody0 = 1
            .NUMbody1 = k
            .body0_pos = Makever(4.5 * Cos(A), 4.5 * Sin(A))
            .body1_pos = Makever(0, 0)
            .Friction_Attenuation = 0.1
        End With
      Next

    '//////////////////////////////////////////////////////////
    For k = 22 To 31
        A = PITWO * (k - 22) / 10

        box(k).m = 0.2
        Call CREATCIRCLE(k, 0.5) '10根主杆上的球
        box(k).BColor = RGB(72, 72, 0)
                box(k).pos = Add(box(1).pos, Makever(8 * Cos(A), 8 * Sin(A)))
        With Joint(k)
            .NUMbody0 = k - 20
            .NUMbody1 = k
            .body0_pos = Makever(3, 0)
            .body1_pos = Makever(0, 0)
             .Friction_Attenuation = 0.1
        End With
    Next

    For k = 1 To 33
        For S = k + 1 To 34
            COLMAP(k, S) = 1
        Next
    Next

    For k = 2 To 20
        For S = k + 1 To 21
            COLMAP(k, S) = 0
        Next
    Next

    '  For k = 2 To 11
    '   COLMAP(k, k + 20) = 0
    '   COLMAP(k, k + 20 + 1) = 0
    'Next

End Sub

'机构
Public Sub Scene2()
    Dim k As Long, S As Long

    Rclear
    NUMSLOT = 12
    NUMJoint = 12
    NUMBox = 14
    NUM_WALL = 1
    Call Rrsort

    For k = 1 To 12
        Call CREATTriangle(k, 2, 2, 2)
        box(k).pos = Makever(10 - k * 0.1, 13 - k * 1)
        box(k).BColor = &H999999
    Next



    '连接铰链
    For k = 1 To 11
        With Slots(k)
            .NUMbody0 = k
            .NUMbody1 = k + 1
            .body0_0pos = Makever(0, -0.3)
            .body0_1pos = Makever(0, 0.6)
            .body1_pos = Makever(0, 1)
            .Restrictions = 1
        End With
              COLMAP(k, k + 1) = 1 '取消相邻铰链的碰撞关系
    Next

    Make_Wall 1, -2, -1, 21, -1
End Sub

'机器人
Public Sub Scene3()
    Dim k As Long, S As Long

    Rclear
    NUMBox = 7
    NUMTwMotor = 2
    NUMJoint = 6
    NUM_WALL = 1
    NUMSLOT = 5
    Call Rrsort
    Make_Wall 1, 0, -1, 21, -1
    ' Rigid_body 9
    Call CREATBOX(1, 8.2, 4)
    box(1).pos = Makever(14, 6)

    Call CREATBOX(2, 2, 2)
    ' Call CREATCIRCLE(2, 1.1)
    box(2).pos = box(1).pos
    box(2).BColor = &H777777

    Call CREATBOX(3, 5, 1.5)
    box(3).pos = box(1).pos
    box(3).BColor = &H777777
    Call CREATBOX(4, 5, 1.5)
    box(4).pos = box(1).pos
    box(4).BColor = &H777777
    '////////////////////////////////////////////

    Call CREATBOX(5, 2, 2)
    ' Call CREATCIRCLE(5, 1.1)
    box(5).pos = box(1).pos
    box(5).BColor = &H777777

    Call CREATBOX(6, 5, 1.5)
    box(6).pos = box(1).pos
    box(6).BColor = &H609F60
    Call CREATBOX(7, 5, 1.5)
    box(7).pos = box(1).pos
    box(7).BColor = &H609F60

    '连接2
    With Joint(1)
        .NUMbody0 = 1
        .NUMbody1 = 2
        .body0_pos = Makever(-3, 0.9)
        .body1_pos = Makever(0, 0)
    End With

    '连接3
    With Joint(2)
        .NUMbody0 = 2
        .NUMbody1 = 3
        .body0_pos = Makever(0, 0.3)
        .body1_pos = Makever(2, 0)
    End With

    With Joint(3)
        .NUMbody0 = 2
        .NUMbody1 = 4
        .body0_pos = Makever(0, -0.3)
        .body1_pos = Makever(2, 0)
    End With

    With Slots(1)
        .NUMbody0 = 3
        .NUMbody1 = 1
        .body0_0pos = Makever(-1, 0)
        .body0_1pos = Makever(1, 0)
        .body1_pos = Makever(-2.6, -2)
        '.Restrictions = True
    End With

    With Slots(2)
        .NUMbody0 = 4
        .NUMbody1 = 1
        .body0_0pos = Makever(-1, 0)
        .body0_1pos = Makever(1, 0)
        .body1_pos = Makever(-3, -2)
        '.Restrictions = True
    End With

    '//////////////////////////////////////
    '连接2
    With Joint(4)
        .NUMbody0 = 1
        .NUMbody1 = 5
        .body0_pos = Makever(3, 0.9)
        .body1_pos = Makever(0, 0)
    End With

    '连接3
    With Joint(5)
        .NUMbody0 = 5
        .NUMbody1 = 6
        .body0_pos = Makever(0, 0.3)
        .body1_pos = Makever(-2, 0)
    End With

    With Joint(6)
        .NUMbody0 = 5
        .NUMbody1 = 7
        .body0_pos = Makever(0, -0.3)
        .body1_pos = Makever(-2, 0)
    End With

    With Slots(3)
        .NUMbody0 = 6
        .NUMbody1 = 1
        .body0_0pos = Makever(-2, 0)
        .body0_1pos = Makever(2, 0)
        .body1_pos = Makever(2.6, -2)
        '.Restrictions = True
    End With

    With Slots(4)
        .NUMbody0 = 7
        .NUMbody1 = 1
        .body0_0pos = Makever(-2, 0)
        .body0_1pos = Makever(2, 0)
        .body1_pos = Makever(3, -2)
        '.Restrictions = True
    End With

    COLMAP(1, 2) = 1
    COLMAP(2, 3) = 1
    COLMAP(1, 3) = 1
    '取消各部分之间碰撞关系
    For k = 1 To 6
        For S = k + 1 To 7
            COLMAP(k, S) = 1
        Next
    Next

    With TwMotor(1)
        .NUMbody0 = 1
        .NUMbody1 = 2
        .max_angv = 0.04
        .TOR = 0.17
    End With

    With TwMotor(2)
        .NUMbody0 = 1
        .NUMbody1 = 5
        .max_angv = 0.04
        .TOR = 0.17
    End With

End Sub

'机器人2
Public Sub Scene4()
    Dim k As Long, S As Long
    Rclear
    NUMBox = 7
    NUM_WALL = 1
    NUMLinSpring = 1
    NUMSLOT = 5
    NUMTwMotor = 2
    NUMJoint = 6
    Call Rrsort

    Make_Wall 1, -2, -1, 21, -1

    Call CREATBOX(1, 2, 4)
    box(1).pos = Makever(14, 7)

    Call CREATTriangle(2, 4, 2.5, 2.5)

    '     box(2).ang = pi
    box(2).pos = Makever(14, 5)

    COLMAP(1, 2) = 1

    With LinSprings(1)
        .body0_pos = Makever(0, -1)
        .body1_pos = Makever(0, 1)
        .NUMbody0 = 1
        .NUMbody1 = 2
        .FreeLong = 0.1
        .k = 0.1
    End With

    With Slots(1)
        .NUMbody0 = 1
        .NUMbody1 = 2
        .body0_0pos = Makever(0, -1)
        .body0_1pos = Makever(0, 1)
        .body1_pos = Makever(0, -0.5)
        '.Restrictions = True
    End With

    With Slots(2)
        .NUMbody0 = 1
        .NUMbody1 = 2
        .body0_0pos = Makever(0, -1)
        .body0_1pos = Makever(0, 1)
        .body1_pos = Makever(0, 0.5)
        '.Restrictions = True
    End With

    Call CREATTriangle(3, 2.3, 2, 2)
    box(3).ang = PI_DIV2
    box(3).pos = Makever(16, 8)

    Call CREATTriangle(4, 2.3, 2, 2)
    box(4).ang = -PI_DIV2
    box(4).pos = Makever(12, 8)

    With Joint(1)
        .NUMbody0 = 1
        .NUMbody1 = 3
        .body0_pos = Makever(1, 1.6)
        .body1_pos = Makever(0, 0.7)
    End With

            With Joint(2)
            .NUMbody0 = 1
            .NUMbody1 = 4
            .body0_pos = Makever(-1, 1.6)
            .body1_pos = Makever(0, 0.7)
        End With

    For k = 1 To 6
        For S = k + 1 To 7
            COLMAP(k, S) = 1
        Next
    Next

    '//////////////////////////////////////
    Call CREATBOX(5, 3, 3)
    box(5).pos = Makever(11, 7)

    Call CREATBOX(6, 3, 3)
    box(6).pos = Makever(11, 7)

    With TwMotor(1)
        .NUMbody0 = 1
        .NUMbody1 = 3
        .max_angv = 0.1
        .TOR = 0.3
    End With

        With TwMotor(2)
            .NUMbody0 = 1
            .NUMbody1 = 4
            .max_angv = 0.1
            .TOR = -0.3
        End With
    
End Sub

Public Sub show_debug(ParamArray p())
    Dim i As Long
    For i = 0 To UBound(p)
        FormDEBUG.Text1 = FormDEBUG.Text1 & p(i) & "  "
    Next
    FormDEBUG.Text1 = FormDEBUG.Text1 & vbCrLf
End Sub
