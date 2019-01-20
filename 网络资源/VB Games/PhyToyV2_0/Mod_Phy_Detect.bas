Attribute VB_Name = "Mod_Const_Cale"
Option Explicit

'各种连接，铰接计算模块

'处理线性弹簧
Public Sub ProcessLinSpring(Boxa As Rig_Body, Boxb As Rig_Body, body0_pos As D3DVECTOR, body1_pos As D3DVECTOR, FreeLong As Single, k As Single)
    Dim Sprforce As D3DVECTOR, IncreaseLong As Single
    Dim world_pos0 As D3DVECTOR, world_pos1 As D3DVECTOR
    Dim deviation As D3DVECTOR

    world_pos0 = ProjectionBodyPnt_World(body0_pos, Boxa)
    world_pos1 = ProjectionBodyPnt_World(body1_pos, Boxb)

    deviation = Subtract(world_pos0, world_pos1)
    IncreaseLong = MaxVel(VLength(deviation) - FreeLong, 0)

    Sprforce = VScale(Normalize(deviation), IncreaseLong * k)

    Form1.DrawLine world_pos0, world_pos1, &HBB

    applyimpulse Boxa, VScale(Sprforce, -1), world_pos0  '应用正冲量
    applyimpulse Boxb, Sprforce, world_pos1 '应用正冲量
End Sub

'计算扭簧
Public Sub ProcessTwSpring(Boxa As Rig_Body, Boxb As Rig_Body, BaseAng As Single, k As Single)
    Dim Torque As Single, curAng As Single

    curAng = BaseAng + Boxa.ang - Boxb.ang

    Torque = k * curAng

    apply_torque Boxa, -Torque
    apply_torque Boxb, Torque
End Sub

Public Sub ProcessTwMotor(Boxa As Rig_Body, Boxb As Rig_Body, max_angv As Single, TOR As Single)
    Dim Torque As Single, dvAng As Single

    '_____
    '     \
    '       \
    '         \_________
    '     0  max_angv

    If TOR > 0 Then

        dvAng = (Boxa.Vang.z - Boxb.Vang.z)
        If dvAng < max_angv And dvAng > 0 Then
            Torque = TOR * (max_angv - dvAng) / max_angv
        ElseIf dvAng <= 0 Then
            Torque = TOR
        Else
            Torque = 0 'dvAng > max_angv
        End If
        apply_torque Boxa, Torque
        apply_torque Boxb, -Torque

    Else

        dvAng = (Boxb.Vang.z - Boxa.Vang.z)
        If dvAng < max_angv And dvAng > 0 Then
            Torque = -TOR * (max_angv - dvAng) / max_angv
        ElseIf dvAng <= 0 Then
            Torque = -TOR
        Else
            Torque = 0
        End If
        apply_torque Boxa, -Torque
        apply_torque Boxb, Torque

    End If
End Sub
'box (TwMotor(I).NUMbody0), box(TwMotor(I).NUMbody1), TwMotor(I), TwMotor(I).

'计算刚体与墙面的铰接
Public Sub ProcessConstraint_world(Boxa As Rig_Body, body0_pos As D3DVECTOR, body1_pos As D3DVECTOR, Optional Friction_Atte As Single = 0)
    Dim cl As collision
    cl.Ra = RotateZ(body0_pos, Boxa.ang)

    Dim world_pos0 As D3DVECTOR, world_pos1 As D3DVECTOR
    Dim deviation As D3DVECTOR
    world_pos0 = Add(cl.Ra, Boxa.pos)
    world_pos1 = body1_pos

    deviation = Subtract(world_pos0, world_pos1)

    Dim deviation_amount As Single, m_vr_extra As D3DVECTOR
    deviation_amount = VLength(deviation)
    m_vr_extra = VScale(deviation, 0.4)

    '''''''''''''''''''''''''''''''''''''''''''''''''
    cl.Va = Add(cross(Boxa.Vang, cl.Ra), Boxa.v) '   v'= v +ω×R' Va

    Dim Vr As D3DVECTOR
    Vr = Add(m_vr_extra, cl.Va)

    cl.Vn = VLength(Vr)
    If (cl.Vn = 0) Then Exit Sub

    cl.N = VScale(Vr, 1 / cl.Vn)
    cl.Vn2 = -cl.Vn

    cl.Ca = 1 / Boxa.m + Dot(cross(VScale(cross(cl.Ra, cl.N), 1 / Boxa.Iz), cl.Ra), cl.N)

    Dim denominator As Single
    denominator = cl.Ca
    If (denominator < 0.0001) Then Exit Sub
    cl.Pn = VScale(cl.N, (cl.Vn2 / denominator))

    applyimpulse Boxa, cl.Pn, world_pos0  '应用正冲量

    If Friction_Atte <> 0 Then
        Dim Torque As Single
        Torque = Friction_Atte * VLength(cl.Pn)
        If Boxa.Vang.z > 0 Then
            apply_torque Boxa, -Torque
        Else
            apply_torque Boxa, Torque
        End If
    End If

    Form1.Picture1.Circle (world_pos0.X, world_pos0.Y), 0.11, &HFF
    Form1.Picture1.Circle (world_pos1.X, world_pos1.Y), 0.11, &HFF
End Sub

'计算两刚体铰接
Public Sub ProcessConstraint(Boxa As Rig_Body, Boxb As Rig_Body, body0_pos As D3DVECTOR, body1_pos As D3DVECTOR, Optional Friction_Atte As Single = 0)
    Dim cl As collision
    cl.Ra = RotateZ(body0_pos, Boxa.ang)
    cl.Rb = RotateZ(body1_pos, Boxb.ang)

    Dim world_pos0 As D3DVECTOR, world_pos1 As D3DVECTOR
    Dim m_world_pos As D3DVECTOR, deviation As D3DVECTOR
    world_pos0 = Add(cl.Ra, Boxa.pos)
    world_pos1 = Add(cl.Rb, Boxb.pos)
    ' m_world_pos = VScale(Add(world_pos0, world_pos1), 0.5)
    deviation = Subtract(world_pos0, world_pos1)

    Dim deviation_amount As Single, m_vr_extra As D3DVECTOR
    deviation_amount = VLength(deviation)
    m_vr_extra = VScale(deviation, 0.3)

    '''''''''''''''''''''''''''''''''''''''''''''''''
    cl.Va = Add(cross(Boxa.Vang, cl.Ra), Boxa.v) '   v'= v +ω×R' Va
    cl.Vb = Add(cross(Boxb.Vang, cl.Rb), Boxb.v)   ' Vb

    Dim Vr As D3DVECTOR
    Vr = Add(m_vr_extra, Subtract(cl.Va, cl.Vb))

    cl.Vn = VLength(Vr)
    If (cl.Vn = 0) Then Exit Sub

    cl.N = VScale(Vr, 1 / cl.Vn)
    cl.Vn2 = -cl.Vn

    cl.Ca = 1 / Boxa.m + Dot(cross(VScale(cross(cl.Ra, cl.N), 1 / Boxa.Iz), cl.Ra), cl.N)
    cl.Cb = 1 / Boxb.m + Dot(cross(VScale(cross(cl.Rb, cl.N), 1 / Boxb.Iz), cl.Rb), cl.N)

    Dim denominator As Single

    cl.c = cl.Ca + cl.Cb
    If (cl.c < 0.0001) Then Exit Sub
    cl.Pn = VScale(cl.N, (cl.Vn2 / cl.c))

    applyimpulse Boxa, cl.Pn, world_pos0  '应用正冲量
    applyimpulse Boxb, VScale(cl.Pn, -1), world_pos1

    If Friction_Atte <> 0 Then
        Dim Torque As Single
        Torque = Friction_Atte * VLength(cl.Pn)
        If (Boxa.Vang.z - Boxb.Vang.z) > 0 Then
            apply_torque Boxa, -Torque
            apply_torque Boxb, Torque
        Else
            apply_torque Boxa, Torque
            apply_torque Boxb, -Torque
        End If
    End If

    Form1.DrawCircle world_pos0, 0.11, &HFF
    Form1.DrawCircle world_pos1, 0.11, &HFF
End Sub

'计算直槽约束
Public Sub ProcessSLOT(Boxa As Rig_Body, Boxb As Rig_Body, body0_0pos As D3DVECTOR, body0_1pos As D3DVECTOR, body1_pos As D3DVECTOR, Optional Restrictions As Boolean = 1)
    Dim cl As collision
    cl.Rb = RotateZ(body1_pos, Boxb.ang)

    Dim world_pos0 As D3DVECTOR, world_pos1 As D3DVECTOR
    world_pos1 = Add(cl.Rb, Boxb.pos)

    Dim LI As Cline
    LI.starpnt = ProjectionBodyPnt_World(body0_0pos, Boxa)
    LI.endpnt = ProjectionBodyPnt_World(body0_1pos, Boxa)
    CaleLineN LI.starpnt, LI.endpnt, LI.nor, LI.d

    Form1.DrawLine LI.starpnt, LI.endpnt, &H777777
    Form1.DrawCircle world_pos1, 0.11, &H777777

    Dim CZ As D3DVECTOR '垂直点
    Dim k As Single
    Dim Vtra As D3DVECTOR, Extra As Single
    Extra = Dot(world_pos1, LI.nor) - LI.d
    CZ = Add(world_pos1, VScale(LI.nor, -Extra)) '垂直点

    If Restrictions Then '限制在两端点内
        If LI.starpnt.X <> LI.endpnt.X Then
            k = (CZ.X - LI.endpnt.X) / (LI.starpnt.X - LI.endpnt.X)
        Else
            k = (CZ.Y - LI.endpnt.Y) / (LI.starpnt.Y - LI.endpnt.Y)
        End If

        If k < 0 Then
            ProcessConstraint Boxa, Boxb, body0_1pos, body1_pos
            Exit Sub
        ElseIf k > 1 Then
            ProcessConstraint Boxa, Boxb, body0_0pos, body1_pos
            Exit Sub
        End If
    End If

    world_pos0 = CZ
    cl.Ra = Subtract(CZ, Boxa.pos)
    cl.N = LI.nor
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    cl.Va = Add(cross(Boxa.Vang, cl.Ra), Boxa.v) '   v'= v +ω×R
    cl.Vb = Add(cross(Boxb.Vang, cl.Rb), Boxb.v)
    cl.Vn = Dot(cl.N, cl.Va) - Dot(cl.N, cl.Vb) - 0.2 * Extra

    cl.Ca = 1 / Boxa.m + Dot(cross(VScale(cross(cl.Ra, cl.N), 1 / Boxa.Iz), cl.Ra), cl.N)
    cl.Cb = 1 / Boxb.m + Dot(cross(VScale(cross(cl.Rb, cl.N), 1 / Boxb.Iz), cl.Rb), cl.N)
    cl.Pn = VScale(cl.N, ((-cl.Vn) / (cl.Ca + cl.Cb)))
    applyimpulse Boxa, cl.Pn, world_pos0  '应用正冲量
    applyimpulse Boxb, VScale(cl.Pn, -1), world_pos1

End Sub

Public Sub process_collision(Boxa As Rig_Body, Boxb As Rig_Body, collpos As D3DVECTOR, colN As D3DVECTOR, Extra As Single)
    '开始正冲量的计算
    Dim cl As collision
    cl.N = colN
    cl.Ra = Subtract(collpos, Boxa.pos)
    cl.Rb = Subtract(collpos, Boxb.pos)
    cl.Va = Add(cross(Boxa.Vang, cl.Ra), Boxa.v) '   v'= v +ω×R
    cl.Vb = Add(cross(Boxb.Vang, cl.Rb), Boxb.v)
    cl.Vn = Dot(cl.N, cl.Va) - Dot(cl.N, cl.Vb) - 0.2 * Extra
    cl.Vn2 = -Def_Elasticity * cl.Vn '垂直相对速度(碰撞后),0.4是反弹系数
    If cl.Vn >= 0 Then Exit Sub
    '  Ca = 1/m+((Ra×N)/Iz×Ra)N     Cb = 1/m+((Rb×N)/Iz×Rb)N
    cl.Ca = 1 / Boxa.m + Dot(cross(VScale(cross(cl.Ra, cl.N), 1 / Boxa.Iz), cl.Ra), cl.N)
    cl.Cb = 1 / Boxb.m + Dot(cross(VScale(cross(cl.Rb, cl.N), 1 / Boxb.Iz), cl.Rb), cl.N)
    cl.Pn = VScale(cl.N, ((cl.Vn2 - cl.Vn) / (cl.Ca + cl.Cb)))
    applyimpulse Boxa, cl.Pn, collpos '应用正冲量
    applyimpulse Boxb, VScale(cl.Pn, -1), collpos
    Form1.DrawLine collpos, Add(collpos, VScale(cl.Pn, 50)), RGB(255, 0, 0)  '正冲量红色显示

    '开始摩擦冲量的计算
    Dim clnew As collision '已经改变的碰撞
    Dim Vr_all As D3DVECTOR
    clnew.N = cl.N
    clnew.Ra = cl.Ra: clnew.Rb = cl.Rb
    clnew.Va = Add(cross(Boxa.Vang, clnew.Ra), Boxa.v) '   v'= v +ω×R
    clnew.Vb = Add(cross(Boxb.Vang, clnew.Rb), Boxb.v)
    Vr_all = Subtract(clnew.Va, clnew.Vb)

    clnew.tangent_vel = Subtract(Vr_all, VScale(clnew.N, Dot(Vr_all, clnew.N)))  'tangent_vel=Va+Vb-((Va+Vb)·N)*N
    clnew.tangent_speed = VLength(clnew.tangent_vel)
    If clnew.tangent_speed > 0 Then
        Dim t As D3DVECTOR
        D3DXVec3Scale t, clnew.tangent_vel, -1 / clnew.tangent_speed
        'Ca = 1/m+((Ra×T)/Iz×Ra)T
        clnew.Ca = 1 / Boxa.m + D3DXVec3Dot(t, cross(VScale(cross(clnew.Ra, t), 1 / Boxa.Iz), clnew.Ra))
        clnew.Cb = 1 / Boxb.m + D3DXVec3Dot(t, cross(VScale(cross(clnew.Rb, t), 1 / Boxb.Iz), clnew.Ra))
        clnew.c = clnew.Ca + clnew.Cb

        If clnew.c > 0 Then
            '动摩擦系数CUR_Friction
            Dim Ptt As Single  '临时摩擦冲量
            Ptt = clnew.tangent_speed / clnew.c
            If Ptt < CUR_Friction * VLength(cl.Pn) Then '动静摩擦判断
                clnew.Pt = Ptt
            Else
                clnew.Pt = CUR_Friction * VLength(cl.Pn)
            End If
            applyimpulse Boxa, VScale(t, clnew.Pt), collpos '应用摩擦冲量
            applyimpulse Boxb, VScale(t, -clnew.Pt), collpos
            Form1.DrawLine collpos, Add(collpos, VScale(t, 50 * clnew.Pt)), RGB(0, 200, 0)  '摩擦力绿色显示

        End If
    End If

End Sub

Public Sub process_collision_wall(Boxa As Rig_Body, collpos As D3DVECTOR, colN As D3DVECTOR, Extra As Single)

    '刺穿分离
    '    Boxa.pos = Add(VScale(colN, Extra), Boxa.pos)

    '开始正冲量的计算
    Dim cl As collision
    Dim posa As D3DVECTOR
    Dim Vanga As D3DVECTOR
    Dim Va As D3DVECTOR
    posa = Boxa.pos: Vanga = Boxa.Vang:  Va = Boxa.v

    cl.N = colN
    cl.Ra = Subtract(collpos, posa)
    cl.Va = Add(cross(Vanga, cl.Ra), Va) '   v'= v +ω×R
    cl.Vn = Dot(cl.N, cl.Va) - 0.1 * Extra '相对垂直速度
    cl.Vn2 = -Def_Elasticity * cl.Vn '垂直相对速度(碰撞后),Elasticity是反弹系数
    If cl.Vn >= 0 Then Exit Sub

    cl.Ca = 1 / Boxa.m + Dot(cross(VScale(cross(cl.Ra, cl.N), 1 / Boxa.Iz), cl.Ra), cl.N)
    D3DXVec3Scale cl.Pn, cl.N, ((cl.Vn2 - cl.Vn) / cl.Ca)
    applyimpulse Boxa, cl.Pn, collpos  '应用正冲量
    Form1.DrawLine collpos, Add(collpos, VScale(cl.Pn, 50)), RGB(255, 0, 0)  '正冲量红色显示

    '开始摩擦冲量的计算
    Dim clnew As collision '已经改变的碰撞
    clnew.N = cl.N
    clnew.Ra = cl.Ra
    clnew.Va = Add(cross(Vanga, clnew.Ra), Va) '   v'= v +ω×R
    clnew.tangent_vel = Subtract(clnew.Va, VScale(clnew.N, Dot(clnew.Va, clnew.N)))  'tangent_vel=Va+Vb-((Va+Vb)·N)*N
    clnew.tangent_speed = VLength(clnew.tangent_vel)
    If clnew.tangent_speed > 0 Then
        Dim t As D3DVECTOR
        D3DXVec3Scale t, clnew.tangent_vel, -1 / clnew.tangent_speed
        'Ca = 1/m+((Ra×T)/Iz×Ra)T
        clnew.c = 1 / Boxa.m + D3DXVec3Dot(t, cross(VScale(cross(clnew.Ra, t), 1 / Boxa.Iz), clnew.Ra))

        If clnew.c > 0 Then
            '动摩擦系数 CUR_Friction
            Dim Ptt As Single  '临时摩擦冲量
            Ptt = clnew.tangent_speed / clnew.c
            If Ptt < CUR_Friction * VLength(cl.Pn) Then '动静摩擦判断
                clnew.Pt = Ptt '静摩擦
            Else
                clnew.Pt = CUR_Friction * VLength(cl.Pn) '动摩擦
            End If
            applyimpulse Boxa, VScale(t, clnew.Pt), collpos '应用摩擦冲量
            Form1.DrawLine collpos, Add(collpos, VScale(t, 50 * clnew.Pt)), RGB(0, 200, 0)  '摩擦力绿色显示
        End If
    End If

End Sub
