Attribute VB_Name = "mFish"
Option Explicit



Public Type tFish
    Pos           As tVec3
    Vel           As tVec3
    Acc           As tVec3
    R             As Double
    G             As Double
    B             As Double
    BehaveTYPE    As Long

    SumALA        As tVec3
    CountAla      As Double

    SumSEP        As tVec3
    CountSEP      As Double

    SumCOH        As tVec3
    CountCOH      As Double

End Type

Public BBuf       As cCairoSurface

Private Const SideSize As Double = 1024    '512
Public Const SideHalf As Double = SideSize / 2    'let's work on a fixed (squared) SideSize here (512x512)
Private Const InvSideSize As Double = 1 / SideSize

Public Const Border As Double = SideHalf * 0.12

Public Const MaxForce As Double = 0.25 * 2


Private Fish()    As tFish
Public NF         As Long


Private Const PathLen As Long = 4

Public Const MaxSpeed As Double = 5
Private Const MaxSpeed2 As Double = MaxSpeed * MaxSpeed

Public Const MinSpeed As Double = 1.5
Private Const MinSpeed2 As Double = MinSpeed * MinSpeed

Private OT        As cOCTTree
Attribute OT.VB_VarUserMemId = 1073741828

Public Do3D       As Boolean

Private DrawOrder() As tVec3

Public CamTarget  As Long

Public Function GetFishPos(wF As Long) As tVec3
    GetFishPos = Fish(wF).Pos

End Function

Public Sub InitFishes(N As Long)
    Dim i         As Long
    Dim J         As Long
    Dim K         As Long

    InitBehaviors





    NF = N
    ReDim Fish(NF)
    ReDim DrawOrder(NF)
    For i = 1 To NF
        With Fish(i)
            .Pos.X = Rnd * (SideSize) - SideHalf
            .Pos.Y = Rnd * (SideSize) - SideHalf
            If Do3D Then .Pos.Z = Rnd * (SideSize) - SideHalf

            .Vel.X = (Rnd * 2 - 1)
            .Vel.Y = (Rnd * 2 - 1)
            '.Vel.Z = (Rnd * 2 - 1)


            .BehaveTYPE = Int(Rnd * NofBehaviors) + 1

            ' Don't want to have Many types 2, so:
            If .BehaveTYPE = 2 And Rnd < 0.8 Then .BehaveTYPE = Int(Rnd * NofBehaviors) + 1

            '            Do
            '                .R = Rnd: .G = Rnd: .B = Rnd
            '            Loop While .R + .G + .B < 1.65

            .R = BehaviorTable(.BehaveTYPE, .BehaveTYPE).R + (Rnd * 2 - 1) * 0.07
            .G = BehaviorTable(.BehaveTYPE, .BehaveTYPE).G + (Rnd * 2 - 1) * 0.07
            .B = BehaviorTable(.BehaveTYPE, .BehaveTYPE).B + (Rnd * 2 - 1) * 0.07

            If .R < 0 Then .R = 0
            If .G < 0 Then .G = 0
            If .B < 0 Then .B = 0

        End With
    Next


    Set OT = New cOCTTree




End Sub


Public Sub RedrawOn(CC As cCairoContext)    'the Main-Routine will render the entire scene

'    CC.SelectFont "Courier New", 8, vbWhite
    CC.SetLineWidth 1.25

    PrepareCenteredCoordSystem CC

    ComputeBEHAVIOURSAccelleration
    UpDatePos
    If Do3D Then
        '        DrawFishes3D_Simple CC
        DrawFishes3D CC
    Else
        DrawFishes CC
    End If

End Sub



Private Sub PrepareCenteredCoordSystem(CC As cCairoContext)
    Dim X As Double, Y As Double, dx As Double, dy As Double
    Cairo.CalcAspectFit 1, BBuf.Width, BBuf.Height, X, Y, dx, dy, 64
    With CC
        .TranslateDrawings X, Y    'shift the placement of our square Draw-Area within the potentially "non-square" Form-Area
        .ScaleDrawings dx * InvSideSize, dy * InvSideSize    'adapt the scaling in relation to our fixed area-size
        .TranslateDrawings SideHalf, SideHalf    'shift the Coord-Sys into the center of our Area
        .AntiAlias = CAIRO_ANTIALIAS_GRAY
        .SetLineCap CAIRO_LINE_CAP_ROUND
    End With

End Sub



Private Sub DrawFishes(CC As cCairoContext)
    Dim i         As Long
    Dim J         As Long
    Dim V         As tVec3
    Dim P         As tVec3

    Dim PathVec   As tVec3
    Dim Vel2      As Double
    Dim InvD      As Double

    Dim S         As Double

    CC.Paint 1, Cairo.CreateSolidPattern(0#, 0#, 0.35, 0.4)

    For i = 1 To UBound(Fish)

        With Fish(i)
            CC.SetSourceRGBA .R, .G, .B, 1
            '--
            P = .Pos
            CC.Arc P.X, P.Y, 5
            CC.Fill
            CC.MoveTo P.X, P.Y
            CC.LineTo P.X + .Vel.X * 3, P.Y + .Vel.Y * 3
            CC.Stroke
            '--
        End With
    Next
End Sub



Private Sub DrawFishes3D_Simple(CC As cCairoContext)
    Dim i         As Long
    Dim J         As Long
    Dim V         As tVec3
    Dim P         As tVec3

    Dim PathVec   As tVec3
    Dim Vel2      As Double
    Dim InvD      As Double

    Dim S         As Double

    CC.Paint 1, Cairo.CreateSolidPattern(0#, 0#, 0.35, 0.4)

    '3D cornice
    CC.SetSourceRGBA 1, 1, 1, 0.5
    CC.Rectangle -SideHalf, -SideHalf, SideSize, SideSize
    CC.Stroke
    CC.Save
    S = 0.33 + 0.66 * (0) * InvSideSize
    CC.ScaleDrawings S, S
    CC.Rectangle -SideHalf, -SideHalf, SideSize, SideSize
    CC.Stroke
    CC.Restore
    ''-----------


    For i = 1 To UBound(Fish)

        With Fish(i)


            '3D
            CC.Save
            S = 0.33 + 0.66 * (.Pos.Z + SideHalf) * InvSideSize
            CC.ScaleDrawings S, S
            '--

            CC.SetSourceRGBA .R, .G, .B, 0.25 + 0.75 * (.Pos.Z + SideHalf) * InvSideSize

            P = .Pos

            CC.MoveTo P.X, P.Y - 5
            CC.LineTo P.X, P.Y + 5
            CC.LineTo P.X + .Vel.X * 6, P.Y + .Vel.Y * 6
            CC.Fill

            If .Vel.Z < 0 Then CC.SetSourceRGBA .R * 0.7, .G * 0.7, .B * 0.7, 0.25 + 0.75 * (.Pos.Z + SideHalf) * InvSideSize

            CC.Arc P.X, P.Y, 5
            CC.Fill

            '3D
            CC.Restore
            '--
        End With
    Next
End Sub

Private Sub DrawFishes3D(CC As cCairoContext)
    Dim i         As Long
    Dim J         As Long
    Dim V         As tVec3
    Dim P         As tVec3
    Dim P2        As tVec3
    Dim Rad       As Double
    Dim Vel       As tVec3
    Dim Pos       As tVec3


    Dim PathVec   As tVec3
    Dim Vel2      As Double
    Dim InvD      As Double

    Dim S         As tVec3

    Dim Pc(1 To 8) As tVec3
    Dim R         As Double
    Dim G         As Double
    Dim B         As Double

    Dim PS        As tVec3
    Dim PSt       As tVec3
    Dim PV        As tVec3
    Dim PVt       As tVec3


    If CamTarget Then
        'Camera.cTo = Fish(CamTarget).Pos
        Camera.cTo = vec3SUM(vec3MUL(Camera.cTo, 0.85), vec3MUL(Fish(CamTarget).Pos, 0.15))
        UpdateCamera
    End If

    CC.Paint 1, Cairo.CreateSolidPattern(0#, 0#, 0.35, 0.4)


    '    Standard
    Pc(1) = PointToScreen(Vec3(-SideHalf, -SideHalf, -SideHalf))
    Pc(2) = PointToScreen(Vec3(SideHalf, -SideHalf, -SideHalf))
    Pc(3) = PointToScreen(Vec3(SideHalf, SideHalf, -SideHalf))
    Pc(4) = PointToScreen(Vec3(-SideHalf, SideHalf, -SideHalf))

    Pc(5) = PointToScreen(Vec3(-SideHalf, -SideHalf, SideHalf))
    Pc(6) = PointToScreen(Vec3(SideHalf, -SideHalf, SideHalf))
    Pc(7) = PointToScreen(Vec3(SideHalf, SideHalf, SideHalf))
    Pc(8) = PointToScreen(Vec3(-SideHalf, SideHalf, SideHalf))

    '  by Testing LineToScreen
    '    Pc(1) = (Vec3(-SideHalf, -SideHalf, -SideHalf))
    '    Pc(2) = (Vec3(SideHalf, -SideHalf, -SideHalf))
    '    Pc(3) = (Vec3(SideHalf, SideHalf, -SideHalf))
    '    Pc(4) = (Vec3(-SideHalf, SideHalf, -SideHalf))
    '    Pc(5) = (Vec3(-SideHalf, -SideHalf, SideHalf))
    '    Pc(6) = (Vec3(SideHalf, -SideHalf, SideHalf))
    '    Pc(7) = (Vec3(SideHalf, SideHalf, SideHalf))
    '    Pc(8) = (Vec3(-SideHalf, SideHalf, SideHalf))
    'LineToScreen Pc(1), Pc(5), Pc(1), Pc(5)
    'LineToScreen Pc(2), Pc(6), Pc(2), Pc(6)
    'LineToScreen Pc(3), Pc(7), Pc(3), Pc(7)
    'LineToScreen Pc(4), Pc(8), Pc(4), Pc(8)



    With CC
        .SetSourceRGBA 1, 1, 1, 0.33
        .MoveTo Pc(1).X, Pc(1).Y
        .LineTo Pc(2).X, Pc(2).Y
        .LineTo Pc(3).X, Pc(3).Y
        .LineTo Pc(4).X, Pc(4).Y
        .LineTo Pc(1).X, Pc(1).Y
        '.Stroke
        .MoveTo Pc(5).X, Pc(5).Y
        .LineTo Pc(6).X, Pc(6).Y
        .LineTo Pc(7).X, Pc(7).Y
        .LineTo Pc(8).X, Pc(8).Y
        .LineTo Pc(5).X, Pc(5).Y
        '.Stroke
        .MoveTo Pc(1).X, Pc(1).Y
        .LineTo Pc(5).X, Pc(5).Y
        .MoveTo Pc(2).X, Pc(2).Y
        .LineTo Pc(6).X, Pc(6).Y
        .MoveTo Pc(3).X, Pc(3).Y
        .LineTo Pc(7).X, Pc(7).Y
        .MoveTo Pc(4).X, Pc(4).Y
        .LineTo Pc(8).X, Pc(8).Y
        .Stroke

    End With


    UpDateDrawOrder

    For i = 1 To UBound(Fish)

        With Fish(DrawOrder(i).X)
            R = .R
            G = .G
            B = .B
            Pos = .Pos
            Vel = .Vel
            P = PointToScreen(Pos)
        End With

        With CC


            ''Rad = SideHalf / Sqr(DistFromCamera2(.Pos))
            'Rad = 8 * SideHalf / Sqr(DrawOrder(I).Y)
            Rad = 10 * SideHalf / (P.Z) * Camera.Zoom

            '---------- FLOOR SHADOW ......... do not need draworder
            ''.SetSourceColor 0
            ''PS.X = Pos.X + 300
            ''PS.Y = SideHalf
            ''PS.Z = Pos.Z - 300
            ''PSt = PointToScreen(PS)
            ''            .Arc PSt.X, PSt.Y, Rad
            ''            .Fill
            ''PV = vec3SUM(Pos, vec3MUL(Vel, 5))
            ''PV.X = PV.X + 300
            ''PV.Y = SideHalf
            ''PV.Z = PV.Z - 300
            ''PVt = PointToScreen(PV)
            ''            .LineTo PSt.X, PSt.Y + Rad
            ''            .LineTo PVt.X, PVt.Y
            ''            .LineTo PSt.X, PSt.Y - Rad
            ''            .Fill
            '----------------------------------------


            P2 = PointToScreen(vec3SUM(Pos, vec3MUL(Vel, 5)))

            .SetSourceRGBA R, G, B, 1
            .MoveTo P.X, P.Y - Rad
            .LineTo P.X, P.Y + Rad
            .LineTo P2.X, P2.Y
            .Fill

            .SetSourceRGB 0, 0, 0    ' Triangle contour
            .LineTo P.X, P.Y + Rad
            .LineTo P2.X, P2.Y
            .LineTo P.X, P.Y - Rad
            .Stroke

            If vec3DOT(vec3SUB(Camera.cTo, Camera.cFrom), Vel) > 0 Then
                .SetSourceRGBA R * 0.75, G * 0.75, B * 0.75, 1
            Else
                .SetSourceRGBA R, G, B, 1
            End If

            .Arc P.X, P.Y, Rad
            .Fill



            'PS.Z = vec3DOT(P, Vec3(0, 1, 0))
            'PS.X = vec3DOT(P, Vec3Cross(Vec3(0, -1, 0), Vec3(0, 0, -1)))
            'PS.Y = SideHalf



        End With

    Next
End Sub



Private Sub UpDatePos()
    Dim Vel2      As Double
    Dim i         As Long
    For i = 1 To UBound(Fish)
        With Fish(i)

            '.Vel = vec3SUM(.Vel, .Acc)

            'Do it more "fishish" --- Attenuate Vertical  (Y)  ACC
            .Vel = vec3SUM(.Vel, Vec3(.Acc.X, .Acc.Y * 0.85, .Acc.Z))


            '.Acc = Vec3(0, 0, 0)'-----<<<<<
            .Acc = vec3MUL(.Acc, 0.75)    '-----<<<<<

            'Limit VEL--------------
            .Vel = vec3Limit(.Vel, MaxSpeed)
            .Vel = vec3LimitMIN(.Vel, MinSpeed)
            '-----------------------
            ' MOVE--------
            .Pos = vec3SUM(.Pos, .Vel)
            '------------

            If .Pos.X < -SideHalf + Border Then .Vel.X = .Vel.X + 0.5
            If .Pos.X > SideHalf - Border Then .Vel.X = .Vel.X - 0.5

            If .Pos.Y < -SideHalf + Border Then .Vel.Y = .Vel.Y + 0.5
            If .Pos.Y > SideHalf - Border Then .Vel.Y = .Vel.Y - 0.5

            If .Pos.Z < -SideHalf + Border Then .Vel.Z = .Vel.Z + 0.5
            If .Pos.Z > SideHalf - Border Then .Vel.Z = .Vel.Z - 0.5
        End With
    Next
End Sub

Private Sub ComputeBEHAVIOURSAccelleration()


' MODIFIED VERSION OF
' https://github.com/OwenMcNaughton/Boids.js/blob/master/js/Boid.js


    Dim i As Long, JJ As Long, J As Long
    Dim CubeL     As tVec3
    Dim CubeH     As tVec3
    Dim rX() As Double, rY() As Double, rZ() As Double, rIDX() As Long
    Dim dx As Double, dy As Double, Dz As Double, D2 As Double
    Dim R         As Double
    Dim Diam2     As Double
    Dim Pos       As tVec3
    Dim BHTij     As tBehavior
    Dim BHTji     As tBehavior
    Dim Diff      As tVec3
    Dim DOTij     As Double
    Dim DOTji     As Double
    Dim V         As tVec3



    OT.Setup -SideHalf, -SideHalf, -SideHalf, SideHalf, SideHalf, SideHalf, 40

    For i = 1 To UBound(Fish)
        With Fish(i)
            OT.InsertSinglePoint .Pos.X, .Pos.Y, .Pos.Z, i
        End With
    Next

    R = MaxInteractDist
    '    Diam2 = (R * 2) * (R * 2) '''<<< WROKNG

    For i = 1 To UBound(Fish)
        With Fish(i)
            Pos = .Pos

            '            CubeL = vec3SUB(Pos, Vec3(R, R, R))
            '            CubeH = vec3SUM(Pos, Vec3(R, R, R))
            '            OT.QueryCube CubeL.X, CubeL.Y, CubeL.Z, _
                         '                         CubeH.X, CubeH.Y, CubeH.Z, _
                         '                         rX(), rY(), rZ(), rIDX()

            OT.QuerySphere Pos.X, Pos.Y, Pos.Z, MaxInteractDist, _
                           rX(), rY(), rZ(), rIDX()

            For JJ = 1 To UBound(rX)
                If i < rIDX(JJ) Then
                    dx = rX(JJ) - Pos.X
                    dy = rY(JJ) - Pos.Y
                    Dz = rZ(JJ) - Pos.Z
                    D2 = dx * dx + dy * dy + Dz * Dz

                    '                    If D2 <= MaxInteractDist2 Then  ' No need for query-Sphere

                    J = rIDX(JJ)

                    DOTij = vec3DOT(.Vel, Vec3(dx, dy, Dz))
                    DOTji = vec3DOT(Fish(J).Vel, Vec3(-dx, -dy, -Dz))


                    BHTij = BehaviorTable(.BehaveTYPE, Fish(J).BehaveTYPE)
                    BHTji = BehaviorTable(Fish(J).BehaveTYPE, .BehaveTYPE)

                    If DOTij > -50 Then
                        'Alignment i,j
                        If D2 < BHTij.Distances.Alignment Then
                            .SumALA = vec3SUM(.SumALA, vec3MUL(Fish(J).Vel, BHTij.Strength.Alignment))
                            .CountAla = .CountAla + 1
                        End If

                        'Cohesion i,j
                        If D2 < BHTij.Distances.Cohesion Then

                            ' .SumCOH = vec3SUM(.SumCOH, Fish(J).Pos)
                            Diff = Vec3Normalize(vec3SUB(Fish(J).Pos, .Pos))
                            Diff = vec3MUL(Diff, BHTij.Strength.Cohesion)
                            .SumCOH = vec3SUM(.SumCOH, Diff)
                            .CountCOH = .CountCOH + 1
                        End If

                        'Separation i,j
                        If D2 < BHTij.Distances.Separation Then
                            Diff = Vec3Normalize(vec3SUB(Pos, Fish(J).Pos))
                            'Diff = vec3MUL(Diff, BHTij.Distances.InvSeparation / Sqr(D2))
                            Diff = vec3MUL(Diff, fastEXP(-4 * D2 * BHTij.Distances.InvSeparation))
                            Diff = vec3MUL(Diff, BHTij.Strength.Separation)
                            .SumSEP = vec3SUM(.SumSEP, Diff)
                            .CountSEP = .CountSEP + 1
                        End If
                    End If

                    If DOTji > -50 Then
                        '                                          If Fish(J).BehaveTYPE = 2 And .BehaveTYPE = 3 Then Stop

                        'Alignment j,i
                        If D2 < BHTji.Distances.Alignment Then
                            Fish(J).SumALA = vec3SUM(Fish(J).SumALA, vec3MUL(.Vel, BHTji.Strength.Alignment))
                            Fish(J).CountAla = Fish(J).CountAla + 1
                        End If

                        'Cohesion j,i
                        If D2 < BHTji.Distances.Cohesion Then
                            ' Fish(J).SumCOH = vec3SUM(Fish(J).SumCOH, .Pos)
                            Diff = Vec3Normalize(vec3SUB(.Pos, Fish(J).Pos))
                            Diff = vec3MUL(Diff, BHTji.Strength.Cohesion)
                            Fish(J).SumCOH = vec3SUM(Fish(J).SumCOH, Diff)
                            Fish(J).CountCOH = Fish(J).CountCOH + 1
                        End If

                        'Separation j,i
                        If D2 < BHTji.Distances.Separation Then
                            Diff = Vec3Normalize(vec3SUB(Fish(J).Pos, Pos))
                            'Diff = vec3MUL(Diff, BHTji.Distances.InvSeparation / Sqr(D2))
                            Diff = vec3MUL(Diff, fastEXP(-4 * D2 * BHTji.Distances.InvSeparation))
                            Diff = vec3MUL(Diff, BHTji.Strength.Separation)
                            Fish(J).SumSEP = vec3SUM(Fish(J).SumSEP, Diff)
                            Fish(J).CountSEP = Fish(J).CountSEP + 1
                        End If

                    End If


                    '                    End If


                End If

            Next
        End With
    Next

    For i = 1 To UBound(Fish)
        With Fish(i)

            If .CountAla Then
                .SumALA = vec3MUL(.SumALA, 1 / .CountAla)
                '.SumALA = Vec3Normalize(.SumALA)
                '.SumALA = vec3MUL(.SumALA, MaxSpeed)
                '.SumALA = vec3SUB(.SumALA, .Vel)
                .SumALA = vec3Limit(.SumALA, MaxForce)
            End If

            If .CountCOH Then
                .SumCOH = vec3MUL(.SumCOH, 1 / .CountCOH)
                '.SumCOH = Vec3Normalize(.SumCOH)
                '.SumCOH = vec3MUL(.SumCOH, MaxSpeed)
                '.SumCOH = vec3SUB(.SumCOH, .Vel)
                .SumCOH = vec3Limit(.SumCOH, MaxForce)
            End If

            If .CountSEP Then
                .SumSEP = vec3MUL(.SumSEP, 1 / .CountSEP)
                '.SumSEP = Vec3Normalize(.SumSEP)
                '.SumSEP = vec3MUL(.SumSEP, MaxSpeed)
                '.SumSEP = vec3SUB(.SumSEP, .Vel)
                '.SumSEP = vec3Limit(.SumSEP, MaxForce)
            End If


            .Acc = vec3SUM(.Acc, .SumALA)
            .Acc = vec3SUM(.Acc, .SumCOH)
            .Acc = vec3SUM(.Acc, .SumSEP)

            If Do3D Then
                .Acc = vec3SUM(.Acc, Vec3(0.025 * (Rnd * 2 - 1), _
                                          0.025 * (Rnd * 2 - 1), _
                                          0.025 * (Rnd * 2 - 1)))
            End If

            '            .Acc = vec3Limit(.Acc, MaxForce)


            If .BehaveTYPE = 1 Then    'slightly go to center
                V = Vec3Normalize(.Pos)
                .Acc = vec3SUM(.Acc, vec3MUL(V, -0.05))
            End If


            .CountAla = 0
            .CountCOH = 0
            .CountSEP = 0
            .SumALA = Vec3(0, 0, 0)
            .SumCOH = Vec3(0, 0, 0)
            .SumSEP = Vec3(0, 0, 0)
        End With
    Next


End Sub



Private Sub UpDateDrawOrder()
    Dim i         As Long
    For i = 1 To NF
        DrawOrder(i).X = i
        DrawOrder(i).Y = DistFromCamera2(Fish(i).Pos)
    Next

    QuickSortDescending 1, NF, 0

End Sub


Private Sub QuickSortDescending(ByVal First As Long, ByVal Last As Long, ByRef Level As Long)
    Dim LOW       As Long
    Dim HIGH      As Long
    Dim dblMidValue As Double
    Dim TMP       As tVec3

    LOW = First
    HIGH = Last

    dblMidValue = DrawOrder((First + Last) \ 2).Y
    Do
        While DrawOrder(LOW).Y > dblMidValue
            LOW = LOW + 1
        Wend

        While DrawOrder(HIGH).Y < dblMidValue
            HIGH = HIGH - 1
        Wend
        If LOW <= HIGH Then
            '---- Swap
            TMP = DrawOrder(LOW)
            DrawOrder(LOW) = DrawOrder(HIGH)
            DrawOrder(HIGH) = TMP
            '-------------------
            LOW = LOW + 1
            HIGH = HIGH - 1
        End If
    Loop While LOW < HIGH

    Level = Level + 1

    'If Level < 2 Then
    If First < HIGH Then QuickSortDescending First, HIGH, Level
    If LOW < Last Then QuickSortDescending LOW, Last, Level
    'End If

End Sub

Public Sub Fish3Dto2D()
    Dim i         As Long

    For i = 1 To NF
        Fish(i).Pos.Z = 0
        Fish(i).Vel.Z = 0
        Fish(i).Acc.Z = 0

    Next
End Sub
