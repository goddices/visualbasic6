Attribute VB_Name = "Mod_gameMain"
Option Explicit

'主循环，相交检测模块

Public Sub mainloop()
    Dim Ticks As Long, i As Long, k As Long, S As Long

    Do

        'Me.Caption = "VB GAME DEV"
        Dim Torque As Single
        If (GetKeyState(vbKeyRight) Or GetKeyState(vbKeyUp)) And &HF0000000 Then
            If Abs(box(2).Vang.z < 0.3) Then
                Torque = -(box(2).Vang.z + 0.3) * 0.4
                apply_torque box(2), Torque
                apply_torque box(19), -Torque
            End If
        End If

        If (GetKeyState(vbKeyLeft) Or GetKeyState(vbKeyDown)) And &HF0000000 Then
            If Abs(box(2).Vang.z < 0.3) Then
                Torque = (0.3 - box(2).Vang.z) * 0.4
                apply_torque box(2), Torque
                apply_torque box(19), -Torque
            End If
        End If

        Ticks = GetTickCount()
        Form1.Picture1.Cls
        Movebox

        For i = 1 To NUM_WALL
            DWAW_Wall i
        Next

        With Form1
            '显示鼠标坐标
            .Picture1.CurrentX = -1
            .Picture1.CurrentY = 14.5
            .Picture1.Print "X " & Format(Mouse_pos.X, "###0.000")
            .Picture1.CurrentX = -1
            .Picture1.CurrentY = 14
            .Picture1.Print "Y " & Format(Mouse_pos.Y, "###0.000")
        End With

        '处理抓取器
        If force_CE And MouseHit_Num <> -1 Then
            With MouseJoint
                .NUMbody0 = MouseHit_Num
                .NUMbody1 = -1
                .body0_pos = MouseHit_pos
                .body1_pos = Mouse_pos
                .Friction_Attenuation = 0.3
            End With
            ProcessConstraint_world box(MouseJoint.NUMbody0), MouseJoint.body0_pos, MouseJoint.body1_pos, MouseJoint.Friction_Attenuation

            '绘制抓取器  兰色叉
            Form1.DrawLine Add(Mouse_pos, Makever(-0.5, -0.5)), Add(Mouse_pos, Makever(0.5, 0.5)), RGB(0, 0, 150)
            Form1.DrawLine Add(Mouse_pos, Makever(0.5, -0.5)), Add(Mouse_pos, Makever(-0.5, 0.5)), RGB(0, 0, 150)
        End If

        '处理铰接
        For i = 1 To NUMJoint
            If Joint(i).NUMbody0 = Joint(i).NUMbody1 Then GoTo NOCAL1

            If (Joint(i).NUMbody1) = -1 Then
                ProcessConstraint_world box(Joint(i).NUMbody0), Joint(i).body0_pos, Joint(i).body1_pos, Joint(i).Friction_Attenuation
            Else
                ProcessConstraint box(Joint(i).NUMbody0), box(Joint(i).NUMbody1), Joint(i).body0_pos, Joint(i).body1_pos, Joint(i).Friction_Attenuation
            End If

NOCAL1:
        Next
        
        '处理直槽
        For i = 1 To NUMSLOT
            If Slots(i).NUMbody0 <> Slots(i).NUMbody1 Then
                ProcessSLOT box(Slots(i).NUMbody0), box(Slots(i).NUMbody1), Slots(i).body0_0pos, Slots(i).body0_1pos, Slots(i).body1_pos, Slots(i).Restrictions
            End If
        Next
        
        '处理线性弹簧 橡皮筋
        For i = 1 To NUMLinSpring
            If LinSprings(i).NUMbody0 <> LinSprings(i).NUMbody1 Then
                ProcessLinSpring box(LinSprings(i).NUMbody0), box(LinSprings(i).NUMbody1), LinSprings(i).body0_pos, LinSprings(i).body1_pos, LinSprings(i).FreeLong, LinSprings(i).k
            End If
        Next
        

        '处理扭簧
        For i = 1 To NUMTwSpring
            If TwSpring(i).k <> 0 Then
                ProcessTwSpring box(TwSpring(i).NUMbody0), box(TwSpring(i).NUMbody1), TwSpring(i).BaseAng, TwSpring(i).k
            End If
        Next

        '处理转动电马达
        For i = 1 To NUMTwMotor
            If TwMotor(i).TOR <> 0 Then
                ProcessTwMotor box(TwMotor(i).NUMbody0), box(TwMotor(i).NUMbody1), TwMotor(i).max_angv, TwMotor(i).TOR
            End If
        Next
        '//////////////////////////////////////////////////////////

        '处理刚体与墙的碰撞
        For k = 1 To NUMBox
            CUR_Friction = box(k).Friction
            If box(k).m = 0 Then GoTo NOCLO

            For S = 1 To NUM_WALL
                If box(k).TAPE = TAPECIRCLE Then '球与墙检测
                    coll_detectCircleTOWall k, S

                ElseIf box(k).TAPE = TAPEBOX Then '盒与墙检测
                    If Abs(Dot(box(k).pos, WALL(S).nor) - WALL(S).d) <= box(k).Rbou Then '包围球可作为粗略检测
                        coll_detectBoxTOWall k, S
                    End If
                End If
            Next
NOCLO:
        Next

        '处理刚体间的碰撞
        For k = 1 To NUMBox
            If box(k).m = 0 Then GoTo NOCLO2
            For S = k + 1 To NUMBox
                If (Not COLMAP(k, S)) And box(S).m <> 0 Then
                    If VDst(box(k).pos, box(S).pos) <= (box(k).Rbou + box(S).Rbou) Then  '包围球可作为粗略检测

                        CUR_Friction = (box(k).Friction + box(S).Friction) / 2

                        If (box(k).TAPE = TAPEBOX And box(S).TAPE = TAPEBOX) Then
                            coll_detectBoxTOBox k, S
                            coll_detectBoxTOBox S, k
                        ElseIf (box(k).TAPE = TAPECIRCLE And box(S).TAPE = TAPEBOX) Then
                            coll_detectCircleTOBox k, S
                        ElseIf (box(k).TAPE = TAPEBOX And box(S).TAPE = TAPECIRCLE) Then
                            coll_detectCircleTOBox S, k
                            '球与球检测
                        ElseIf (box(k).TAPE = TAPECIRCLE And box(S).TAPE = TAPECIRCLE) Then
                            coll_detectCirTOCir k, S
                        End If

                    End If

                End If

            Next
NOCLO2:
        Next 'END For k = 1 To NUMBox

        Do
            DoEvents

        Loop Until GetTickCount() - Ticks > 10
    Loop

End Sub

'应用转矩
Public Sub apply_torque(applyBox As Rig_Body, Torque As Single)
    Dim AccRou As D3DVECTOR
    AccRou = Makever(0, 0, Torque / applyBox.Iz)
    D3DXVec3Add applyBox.Vang, applyBox.Vang, AccRou
End Sub

'应用瞬间冲量
Public Sub applyimpulse(applyBox As Rig_Body, impulse As D3DVECTOR, colpos As D3DVECTOR)
    Dim Ra As D3DVECTOR
    D3DXVec3Subtract Ra, colpos, applyBox.pos
    D3DXVec3Add applyBox.v, applyBox.v, VScale(impulse, 1 / applyBox.m)
    D3DXVec3Add applyBox.Vang, applyBox.Vang, VScale(cross(Ra, impulse), 1 / applyBox.Iz)
End Sub

Public Sub Movebox()
    Dim i As Long, k As Long

    For k = 1 To NUMBox
        If box(k).m = 0 Then GoTo NOCLO3

        D3DXVec3Add box(k).pos, box(k).pos, box(k).v
        box(k).ang = box(k).ang + box(k).Vang.z
        box(k).v.Y = box(k).v.Y - 0.00118

        If (box(k).TAPE = TAPEBOX) Then ''''''''''''''''''''''''''''''''''''''''
            Dim nvs As Long
            nvs = box(k).numverts
            For i = 1 To nvs
                box(k).vertes(i) = ProjectionBodyPnt_World(box(k).tvertes(i), box(k))
            Next

            For i = 1 To nvs
                CaleLineN box(k).vertes(i), box(k).vertes((i Mod nvs) + 1), box(k).nor(i), box(k).d(i)
                Form1.DrawLine box(k).vertes(i), box(k).vertes((i Mod nvs) + 1), box(k).BColor
            Next

            ''''''''''''''''''''''''''''''''''''''''
        ElseIf (box(k).TAPE = TAPECIRCLE) Then
            Dim Vc As D3DVECTOR

            Vc = ProjectionBodyPnt_World(Makever(box(k).Rbou, 0), box(k))

            Form1.DrawLine box(k).pos, Vc, box(k).BColor
            Form1.DrawCircle box(k).pos, box(k).Rbou, box(k).BColor
        End If

        Form1.Picture1.CurrentX = box(k).pos.X
        Form1.Picture1.CurrentY = box(k).pos.Y + 0.1
        ' Picture1.ForeColor = &H606060
        Form1.Picture1.Print k
        'Picture1.ForeColor = &O0
'Form1.Picture1.PSet (box(k).pos.X, box(k).pos.Y)
NOCLO3:
    Next
End Sub

Public Sub DWAW_Wall(NUM As Long)
    Form1.DrawLine WALL(NUM).starpnt, WALL(NUM).endpnt
End Sub

'盒子与盒子检测，撞击计算
Public Sub coll_detectBoxTOBox(num1 As Long, num2 As Long)
    Dim d As Single, collnormal As D3DVECTOR
    Dim i As Long, j As Long
    Dim nvs1 As Long, nvs2 As Long
    nvs1 = box(num1).numverts
    nvs2 = box(num2).numverts

    For i = 1 To nvs1
        For j = 1 To nvs2
            d = lineins(box(num1).vertes(i), box(num1).pos, box(num2).vertes(j), box(num2).vertes((j Mod nvs2) + 1), collnormal)

            If d <> 0 Then process_collision box(num1), box(num2), box(num1).vertes(i), collnormal, d

        Next
    Next

End Sub

'球与盒子检测，撞击计算
Public Sub coll_detectCircleTOBox(CirNum As Long, BoxNum As Long)
    Dim i As Long, j As Long, nvs As Long
    Dim tmpLine As Cline, d As Single, collnormal As D3DVECTOR, collpos As D3DVECTOR
    nvs = box(BoxNum).numverts

    For i = 1 To nvs
        tmpLine.starpnt = box(BoxNum).vertes(i)
        tmpLine.endpnt = box(BoxNum).vertes((i Mod nvs) + 1)
        tmpLine.nor = box(BoxNum).nor(i)
        tmpLine.d = box(BoxNum).d(i)
        d = CircleinsLine(tmpLine, box(CirNum).pos, box(CirNum).Rbou, collnormal, collpos)
        If d > 0 Then process_collision box(CirNum), box(BoxNum), collpos, collnormal, d
    Next

    For j = 1 To nvs
        d = CircleinsPnt(box(BoxNum).vertes(j), box(CirNum).pos, box(CirNum).Rbou, collnormal)
        If d > 0 Then process_collision box(CirNum), box(BoxNum), box(BoxNum).vertes(j), collnormal, d
    Next
End Sub

'球与球检测，撞击计算
Public Sub coll_detectCirTOCir(CirNum1 As Long, CirNum2 As Long)
    Dim pos1 As D3DVECTOR, pos2 As D3DVECTOR, Vtra As D3DVECTOR, Extra As Single

    pos1 = box(CirNum1).pos: pos2 = box(CirNum2).pos
    Vtra = Subtract(pos1, pos2)
    Extra = VLength(Vtra) - (box(CirNum1).Rbou + box(CirNum2).Rbou)
    If Extra < 0 Then
        Vtra = Normalize(Vtra)
        process_collision box(CirNum1), box(CirNum2), Add(pos2, VScale(Vtra, box(CirNum2).Rbou)), Vtra, -Extra

    End If
End Sub

'盒子与墙检测，撞击计算
Public Sub coll_detectBoxTOWall(BoxNum As Long, WallNum As Long)
    Dim d As Single, collnormal As D3DVECTOR
    Dim i As Long, nvs As Long
    nvs = box(BoxNum).numverts

    For i = 1 To nvs
        d = lineins(box(BoxNum).vertes(i), box(BoxNum).pos, WALL(WallNum).starpnt, WALL(WallNum).endpnt, collnormal)
        If d > 0 Then
            process_collision_wall box(BoxNum), box(BoxNum).vertes(i), collnormal, d
        ElseIf d < 0 Then
            process_collision_wall box(BoxNum), box(BoxNum).vertes(i), VScale(collnormal, -1), -d '在墙的反面
        End If
    Next

    d = pntINBOX(WALL(WallNum).starpnt, box(BoxNum), collnormal)
    If d > 0 Then
        process_collision_wall box(BoxNum), WALL(WallNum).starpnt, collnormal, d
    End If

    d = pntINBOX(WALL(WallNum).endpnt, box(BoxNum), collnormal)
    If d > 0 Then
        process_collision_wall box(BoxNum), WALL(WallNum).endpnt, collnormal, d
    End If

End Sub

'球与墙检测，撞击计算
Public Sub coll_detectCircleTOWall(CirNum As Long, WallNum As Long)
    Dim d As Single, collnormal As D3DVECTOR, collpos As D3DVECTOR
    d = CircleinsLine(WALL(WallNum), box(CirNum).pos, box(CirNum).Rbou, collnormal, collpos)

    If d <> 0 Then process_collision_wall box(CirNum), collpos, collnormal, d

    '与墙2端点的碰撞
    d = CircleinsPnt(WALL(WallNum).starpnt, box(CirNum).pos, box(CirNum).Rbou, collnormal)
    If d > 0 Then process_collision_wall box(CirNum), WALL(WallNum).starpnt, collnormal, d

    d = CircleinsPnt(WALL(WallNum).endpnt, box(CirNum).pos, box(CirNum).Rbou, collnormal)
    If d > 0 Then process_collision_wall box(CirNum), WALL(WallNum).endpnt, collnormal, d

End Sub

'检测点是否在盒子内，是则返回刺穿距离与撞击点垂直方向
Public Function pntINBOX(pnt As D3DVECTOR, Boxa As Rig_Body, ByRef normal As D3DVECTOR) As Single
    Dim i As Long, nvs As Long
    Dim d() As Single, dmin As Single
    nvs = Boxa.numverts
    ReDim d(nvs)
    
    pntINBOX = 0
    For i = 1 To nvs
        d(i) = Dot(pnt, Boxa.nor(i)) - Boxa.d(i)
        If (d(i) < 0) Then Exit Function
    Next
   
    dmin = 10000

    For i = 1 To nvs
        If d(i) < dmin Then
            dmin = d(i)
            normal = Boxa.nor(i)
        End If
    Next
    pntINBOX = dmin
End Function

'检测点是否在圆内，是则返回刺穿距离与撞击点垂直方向
Public Function CircleinsPnt(pnt As D3DVECTOR, CirP As D3DVECTOR, CirR As Single, ByRef normal As D3DVECTOR) As Single
    Dim Vtra As D3DVECTOR, Extra As Single

    CircleinsPnt = 0
    Vtra = Subtract(CirP, pnt)
    Extra = CirR - VLength(Vtra)
    If Extra > 0 Then
        CircleinsPnt = Extra
        normal = Normalize(Vtra)
    End If

End Function

'检测圆是否与线段相交，是则返回刺穿距离与撞击点垂直方向
Public Function CircleinsLine(LI As Cline, CirP As D3DVECTOR, CirR As Single, ByRef normal As D3DVECTOR, ByRef collpos As D3DVECTOR) As Single
    CircleinsLine = 0
    'Dim nor   As D3DVECTOR, d As Single
    Dim CZ As D3DVECTOR '垂直点
    Dim k As Single '垂直点
    Dim Extra As Single

    Extra = Dot(CirP, LI.nor) - LI.d

    If Abs(Extra) < CirR Then
        CZ = Add(CirP, VScale(LI.nor, -Extra))

        If LI.starpnt.X <> LI.endpnt.X Then
            k = (CZ.X - LI.endpnt.X) / (LI.starpnt.X - LI.endpnt.X)
        Else
            k = (CZ.Y - LI.endpnt.Y) / (LI.starpnt.Y - LI.endpnt.Y)
        End If

        If k > 0 And k < 1 Then
            If Extra > 0 Then
                normal = LI.nor
                CircleinsLine = CirR - Extra
                collpos = Add(CirP, VScale(LI.nor, -CirR))
            Else
                normal = VScale(LI.nor, -1)
                CircleinsLine = CirR + Extra
                collpos = Add(CirP, VScale(LI.nor, CirR))
            End If
        End If 'If k > 0 And k < 1 Then

    End If

End Function

'检测两线段是否相交，是则返回刺穿距离与撞击点垂直方向
Public Function lineins(l11 As D3DVECTOR, l12 As D3DVECTOR, l21 As D3DVECTOR, l22 As D3DVECTOR, ByRef normal As D3DVECTOR) As Single
    lineins = 0
    If cross(Subtract(l22, l11), Subtract(l12, l11)).z * cross(Subtract(l12, l11), Subtract(l21, l11)).z > 0 And _
       cross(Subtract(l12, l21), Subtract(l22, l21)).z * cross(Subtract(l22, l21), Subtract(l11, l21)).z > 0 Then
        normal = Normalize(cross(Subtract(l22, l21), Makever(0, 0, -1)))
        lineins = (Dot(l21, normal) - Dot(l11, normal))
    End If
End Function

'计算由2点构成的线段法线与距离
Public Function CaleLineN(starpnt As D3DVECTOR, endpnt As D3DVECTOR, ByRef normal As D3DVECTOR, ByRef d As Single)
    normal = Normalize(RotateZ(Subtract(starpnt, endpnt), PI_DIV2))
    d = Dot(starpnt, normal)
End Function

'一矢量从刚体自身坐标投影到世界坐标
Public Function ProjectionBodyPnt_World(pnt As D3DVECTOR, Boxa As Rig_Body) As D3DVECTOR
    ProjectionBodyPnt_World = Add(Boxa.pos, RotateZ(pnt, Boxa.ang))
End Function

'一矢量从世界坐标投影到刚体自身坐标
Public Function ProjectionWorldPnt_Body(pnt As D3DVECTOR, Boxa As Rig_Body) As D3DVECTOR
    ProjectionWorldPnt_Body = RotateZ(Subtract(pnt, Boxa.pos), -Boxa.ang)
End Function
