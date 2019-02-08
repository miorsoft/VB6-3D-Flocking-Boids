Attribute VB_Name = "mVectors"
Option Explicit

Public Const PI   As Double = 3.14159265358979
Public Const InvPI As Double = 1 / 3.14159265358979

Public Const PIh  As Double = 1.5707963267949
Public Const PI2  As Double = 6.8318530717959



Private Const INV2 As Double = 0.5
Private Const INV9 As Double = 0.111111111111111
Private Const INV72 As Double = 1.38888888888889E-02
Private Const INV1008 As Double = 9.92063492063492E-04
Private Const INV30240 As Double = 3.30687830687831E-05


Public Type tVec3
    X             As Double
    Y             As Double
    Z             As Double
End Type
Public Function Vec3(X As Double, Y As Double, Z As Double) As tVec3
    Vec3.X = X
    Vec3.Y = Y
    Vec3.Z = Z
End Function
Public Function vec3LEN2(V As tVec3) As Double
    With V
        vec3LEN2 = .X * .X + .Y * .Y + .Z * .Z
    End With
End Function

Public Function Vec3Normalize(V As tVec3) As tVec3
    Dim D         As Double
    With V
        D = .X * .X + .Y * .Y + .Z * .Z
    End With
    If D Then
        D = 1 / Sqr(D)
        Vec3Normalize.X = V.X * D
        Vec3Normalize.Y = V.Y * D
        Vec3Normalize.Z = V.Z * D
    End If
End Function

Public Function vec3SUM(V1 As tVec3, V2 As tVec3) As tVec3
    vec3SUM.X = V1.X + V2.X
    vec3SUM.Y = V1.Y + V2.Y
    vec3SUM.Z = V1.Z + V2.Z
End Function

Public Function vec3SUB(V1 As tVec3, V2 As tVec3) As tVec3
    vec3SUB.X = V1.X - V2.X
    vec3SUB.Y = V1.Y - V2.Y
    vec3SUB.Z = V1.Z - V2.Z
End Function

Public Function vec3MUL(V As tVec3, Scalar As Double) As tVec3
    vec3MUL.X = V.X * Scalar
    vec3MUL.Y = V.Y * Scalar
    vec3MUL.Z = V.Z * Scalar
End Function

Public Function vec3DOT(V1 As tVec3, V2 As tVec3) As Double
    vec3DOT = V1.X * V2.X + V1.Y * V2.Y + V1.Z * V2.Z
End Function

Public Function Vec3Cross(V1 As tVec3, _
                          V2 As tVec3) As tVec3

    Vec3Cross.X = V1.Y * V2.Z - V1.Z * V2.Y
    Vec3Cross.Y = V1.Z * V2.X - V1.X * V2.Z
    Vec3Cross.Z = V1.X * V2.Y - V1.Y * V2.X

End Function

Public Function vec3Limit(V As tVec3, Limit As Double) As tVec3
    Dim D         As Double
    D = vec3LEN2(V)
    If D > Limit * Limit Then
        D = 1 / Sqr(D)
        vec3Limit.X = V.X * D * Limit
        vec3Limit.Y = V.Y * D * Limit
        vec3Limit.Z = V.Z * D * Limit
    Else
        vec3Limit = V
    End If
End Function
Public Function vec3LimitMIN(V As tVec3, Limit As Double) As tVec3
    Dim D         As Double
    D = vec3LEN2(V)
    If D < Limit * Limit Then
        D = 1 / Sqr(D)
        vec3LimitMIN.X = V.X * D * Limit
        vec3LimitMIN.Y = V.Y * D * Limit
        vec3LimitMIN.Z = V.Z * D * Limit
    Else
        vec3LimitMIN = V
    End If
End Function

Public Function vec3Seek(Pos As tVec3, Vel As tVec3, Target As tVec3) As tVec3

    vec3Seek = Vec3Normalize(vec3SUB(Target, Pos))
    vec3Seek = vec3MUL(vec3Seek, MaxSpeed)
    vec3Seek = vec3SUB(vec3Seek, Vel)
    vec3Seek = vec3Limit(vec3Seek, MaxForce)


End Function




Public Function fastEXP(ByVal X As Double) As Double
'https://en.wikipedia.org/wiki/Pad%C3%A9_approximant
    Dim x2        As Double
    Dim X3        As Double
    Dim X4        As Double
    Dim X5        As Double


    If X < 5# Then

        If X < -7# Then fastEXP = 0#: Exit Function

        x2 = X * X
        X3 = x2 * X
        X4 = X3 * X
        X5 = X4 * X

        fastEXP = (1# + INV2 * X + INV9 * x2 + INV72 * X3 + INV1008 * X4 + INV30240 * X5) / _
                  (1# - INV2 * X + INV9 * x2 - INV72 * X3 + INV1008 * X4 - INV30240 * X5)

    Else
        fastEXP = Exp(X)
    End If


End Function

Public Function Atan2(ByVal X As Double, ByVal Y As Double) As Double

'    Atan2 = Cairo.CalcArc(Y, X)
'    End Function

'**********************************************************************************
    If X Then    '''Sempre USATA
        Atan2 = -PI + Atn(Y / X) - (X > 0!) * PI
    Else
        Atan2 = -PIh - (Y > 0!) * PI
    End If
End Function


