Attribute VB_Name = "m3DtoScreen"
Option Explicit

'http://paulbourke.net/geometry/transformationprojection/transform.c

Private Enum eProjection
    PERSPECTIVE
    ORTHOGRAPHIC
End Enum

Private Type tCamera
    cFrom         As tVec3
    cTo           As tVec3
    cUp           As tVec3
    ANGh          As Double
    ANGv          As Double
    Zoom          As Double
    NearPlane     As Double
    FarPlane      As Double
    Projection    As eProjection
End Type

Private Type tScreen
    Center        As tVec3
    Size          As tVec3
End Type

Public Camera     As tCamera
Public Scree      As tScreen

Private Const Epsilon As Double = 0.001
Private Const Deg2Rad As Double = 1.74532925199433E-02     'Degrees to Radians

Private CamNormRightDIR As tVec3
Private CamNormFrontDIR As tVec3
Private CamNormUPDir As tVec3

Private INVtanthetaH As Double
Private INVtanthetaV As Double

Public Sub InitCamera(vFrom As tVec3, _
                      vTo As tVec3)

    With Camera
        .cFrom = vFrom
        .cTo = vTo
        .FarPlane = 999999999
        .NearPlane = 0    ' -99999999
        .cUp.X = 0
        .cUp.Y = -1
        .cUp.Z = 0
        .ANGh = 37
        .ANGv = 37 * Scree.Size.Y / Scree.Size.X
        .Projection = PERSPECTIVE
        .Zoom = 1
        INVtanthetaH = 1 / Tan(.ANGh * Deg2Rad * 0.5)
        INVtanthetaV = 1 / Tan(.ANGv * Deg2Rad * 0.5)
    End With

    UpdateCamera

End Sub
Public Function UpdateCamera()

    CamNormFrontDIR = vec3SUB(Camera.cTo, Camera.cFrom)
    CamNormFrontDIR = Vec3Normalize(CamNormFrontDIR)

    CamNormRightDIR = Vec3Cross(Camera.cUp, CamNormFrontDIR)
    CamNormRightDIR = Vec3Normalize(CamNormRightDIR)

    CamNormUPDir = Vec3Cross(CamNormFrontDIR, CamNormRightDIR)

    '    Calculate camera aperture statics, note: angles in degrees
    '    INVtanthetaH = 1 / Tan(Camera.ANGh * Deg2Rad * 0.5)
    '    INVtanthetaV = 1 / Tan(Camera.ANGv * Deg2Rad * 0.5)

End Function

Public Function DistFromCamera2(V As tVec3) As Double
    Dim dx        As Double
    Dim dy        As Double
    Dim Dz        As Double
    dx = V.X - Camera.cFrom.X
    dy = V.Y - Camera.cFrom.Y
    Dz = V.Z - Camera.cFrom.Z
    DistFromCamera2 = dx * dx + dy * dy + Dz * Dz
End Function


Public Function PointToScreen(P As tVec3) As tVec3
'  ALL HERE
    Dim PO        As tVec3
    Dim EYE       As tVec3
    Dim eyeN      As tVec3
    Dim D         As Double

    '''   Take a point in world coordinates and transform it to
    '''   a point in the eye coordinate system.
    '''   pvWorld2EYE
    PO = vec3SUB(P, Camera.cFrom)
    EYE.X = vec3DOT(PO, CamNormRightDIR)    'Projection
    EYE.Y = vec3DOT(PO, CamNormUPDir)
    EYE.Z = vec3DOT(PO, CamNormFrontDIR)
    '--------------

    ''   Take a vector in eye coordinates and transform it into
    ''   normalised coordinates for a perspective view. No normalisation
    ''   is performed for an orthographic projection. Note that although
    ''   the Z component of the normalised vector is copied from the eye
    ''   coordinate system, it is generally no longer needed. It can
    ''   however still be used externally for vector sorting.
    ''   pvEYE2Norm
    If Camera.Projection = PERSPECTIVE Then
        D = Camera.Zoom / EYE.Z
        eyeN.X = D * EYE.X * INVtanthetaH
        eyeN.Y = D * EYE.Y * INVtanthetaV
        eyeN.Z = EYE.Z
    Else
        eyeN.X = 0.5 * Camera.Zoom * EYE.X * INVtanthetaH / Scree.Size.X
        eyeN.Y = 0.5 * Camera.Zoom * EYE.Y * INVtanthetaV / Scree.Size.Y
        eyeN.Z = EYE.Z
    End If

    '' pvNorm2Screen
    PointToScreen.X = Scree.Center.X + Scree.Size.X * eyeN.X * 0.5    'Not sure about X sign
    PointToScreen.Y = Scree.Center.Y - Scree.Size.Y * eyeN.Y * 0.5

    PointToScreen.Z = eyeN.Z    'dist from camera eye

End Function



Public Sub CameraSetRotation(ByVal Yaw As Double, ByVal Pitch As Double)
    Dim D         As Double
    ' Thanks to Passel:
    ' http://www.vbforums.com/showthread.php?870755-3D-Swimming-Fish-Algorithm&p=5356667&viewfull=1#post5356667

    '    If Pitch > 90 Then Pitch = 90
    '    If Pitch < -90 Then Pitch = -90
    With Camera
        D = Sqr(vec3LEN2(vec3SUB(.cFrom, .cTo)))
        .cFrom.X = .cTo.X + D * (Sin(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
        .cFrom.Y = .cTo.Y + D * (Sin(Pitch * Deg2Rad))
        .cFrom.Z = .cTo.Z + D * (Cos(Yaw * Deg2Rad) * Cos(Pitch * Deg2Rad))
    End With

    UpdateCamera

End Sub









'                     ----------------------------------------------------------------
'                     ----------------------------------------------------------------
'    FOLLOW LINE CLIP ----------------------------------------------------------------
'                     ----------------------------------------------------------------
'                     ----------------------------------------------------------------









Private Function pvClipEYE(ByRef e1 As tVec3, ByRef e2 As tVec3) As Boolean
'   Clip a line segment in eye coordinates to the camera .
'   and back clipping planes. Return FALSE if the line segment
'   is entirely before or after the clipping planes.
    Dim mu        As Double

    ' Is the vector totally in . of the . cutting plane ?
    If (e1.Y <= Camera.NearPlane And e2.Y <= Camera.NearPlane) Then Exit Function

    ' Is the vector totally behind the back cutting plane ?
    If (e1.Y >= Camera.FarPlane And e2.Y >= Camera.FarPlane) Then Exit Function


    ' Is the vector partly in . of the . cutting plane ?
    If ((e1.Y < Camera.NearPlane And e2.Y > Camera.NearPlane) Or _
        (e1.Y > Camera.NearPlane And e2.Y < Camera.NearPlane)) Then
        mu = (Camera.NearPlane - e1.Y) / (e2.Y - e1.Y)
        If (e1.Y < Camera.NearPlane) Then
            e1.X = e1.X + mu * (e2.X - e1.X)
            e1.Z = e1.Z + mu * (e2.Z - e1.Z)
            e1.Y = Camera.NearPlane
        Else
            e2.X = e1.X + mu * (e2.X - e1.X)
            e2.Z = e1.Z + mu * (e2.Z - e1.Z)
            e2.Y = Camera.NearPlane
        End If
    End If

    ' Is the vector partly behind the farplane cutting plane ?
    If ((e1.Y < Camera.FarPlane And e2.Y > Camera.FarPlane) Or _
        (e1.Y > Camera.FarPlane And e2.Y < Camera.FarPlane)) Then
        mu = (Camera.FarPlane - e1.Y) / (e2.Y - e1.Y)
        If (e1.Y < Camera.FarPlane) Then
            e2.X = e1.X + mu * (e2.X - e1.X)
            e2.Z = e1.Z + mu * (e2.Z - e1.Z)
            e2.Y = Camera.FarPlane
        Else
            e1.X = e1.X + mu * (e2.X - e1.X)
            e1.Z = e1.Z + mu * (e2.Z - e1.Z)
            e1.Y = Camera.FarPlane
        End If
    End If

    pvClipEYE = True

End Function

Private Function pvClipNorm(n1 As tVec3, n2 As tVec3) As Boolean
'   Clip a line segment to the normalised coordinate +- square.
'   The y component is not touched.
    Dim mu        As Double

    ' Is the line segment totally right of x = 1 ?
    If (n1.X >= 1 And n2.X >= 1) Then Exit Function

    ' Is the line segment totally left of x = -1 ?
    If (n1.X <= -1 And n2.X <= -1) Then Exit Function

    ' Does the vector cross x = 1 ?
    If ((n1.X > 1 And n2.X < 1) Or (n1.X < 1 And n2.X > 1)) Then
        mu = (1 - n1.X) / (n2.X - n1.X)
        If (n1.X < 1) Then
            n2.Z = n1.Z + mu * (n2.Z - n1.Z)
            n2.X = 1
        Else
            n1.Z = n1.Z + mu * (n2.Z - n1.Z)
            n1.X = 1
        End If
    End If

    ' Does the vector cross x = -1 ?
    If ((n1.X < -1 And n2.X > -1) Or (n1.X > -1 And n2.X < -1)) Then
        mu = (-1 - n1.X) / (n2.X - n1.X)
        If (n1.X > -1) Then
            n2.Z = n1.Z + mu * (n2.Z - n1.Z)
            n2.X = -1
        Else
            n1.Z = n1.Z + mu * (n2.Z - n1.Z)
            n1.X = -1
        End If
    End If

    ' Is the line segment totally above z = 1 ?
    If (n1.Z >= 1 And n2.Z >= 1) Then Exit Function

    ' Is the line segment totally below z = -1 ?
    If (n1.Z <= -1 And n2.Z <= -1) Then Exit Function

    ' Does the vector cross z = 1 ?
    If ((n1.Z > 1 And n2.Z < 1) Or (n1.Z < 1 And n2.Z > 1)) Then
        mu = (1 - n1.Z) / (n2.Z - n1.Z)
        If (n1.Z < 1) Then
            n2.X = n1.X + mu * (n2.X - n1.X)
            n2.Z = 1
        Else
            n1.X = n1.X + mu * (n2.X - n1.X)
            n1.Z = 1
        End If
    End If

    ' Does the vector cross z = -1 ?
    If ((n1.Z < -1 And n2.Z > -1) Or (n1.Z > -1 And n2.Z < -1)) Then
        mu = (-1 - n1.Z) / (n2.Z - n1.Z)
        If (n1.Z > -1) Then
            n2.X = n1.X + mu * (n2.X - n1.X)
            n2.Z = -1
        Else
            n1.X = n1.X + mu * (n2.X - n1.X)
            n1.Z = -1
        End If
    End If

    pvClipNorm = True

End Function

Private Function pvWorld2EYE(P As tVec3) As tVec3
    Dim PO        As tVec3

    '''   Take a point in world coordinates and transform it to
    '''   a point in the eye coordinate system.
    PO = vec3SUB(P, Camera.cFrom)
    pvWorld2EYE.X = vec3DOT(PO, CamNormRightDIR)    'Projection
    pvWorld2EYE.Y = vec3DOT(PO, CamNormUPDir)
    pvWorld2EYE.Z = vec3DOT(PO, CamNormFrontDIR)
    '--------------
End Function
Private Function pvEYE2Norm(EYE As tVec3) As tVec3
''   Take a vector in eye coordinates and transform it into
''   normalised coordinates for a perspective view. No normalisation
''   is performed for an orthographic projection. Note that although
''   the Z component of the normalised vector is copied from the eye
''   coordinate system, it is generally no longer needed. It can
''   however still be used externally for vector sorting.
' Eye to NORM
    Dim D         As Double

    If Camera.Projection = PERSPECTIVE Then
        D = Camera.Zoom / EYE.Z
        pvEYE2Norm.X = D * EYE.X * INVtanthetaH
        pvEYE2Norm.Y = D * EYE.Y * INVtanthetaV
        pvEYE2Norm.Z = EYE.Z
    Else
        pvEYE2Norm.X = 0.5 * Camera.Zoom * EYE.X * INVtanthetaH / Scree.Size.X
        pvEYE2Norm.Y = 0.5 * Camera.Zoom * EYE.Y * INVtanthetaV / Scree.Size.Y
        pvEYE2Norm.Z = EYE.Z
    End If
End Function

Private Function pvNorm2Screen(eyeN As tVec3) As tVec3
    pvNorm2Screen.X = Scree.Center.X + Scree.Size.X * eyeN.X * 0.5    'Not sure about X sign
    pvNorm2Screen.Y = Scree.Center.Y - Scree.Size.Y * eyeN.Y * 0.5
    pvNorm2Screen.Z = eyeN.Z    'dist from camera eye
End Function

Public Function LineToScreen(ByRef W1 As tVec3, ByRef W2 As tVec3, _
                             ByRef P1 As tVec3, ByRef P2 As tVec3) As Boolean

'   Transform and appropriately clip a line segment from
'   world to screen coordinates. Return TRUE if something
'   is visible and needs to be drawn, namely a line between
'   screen coordinates p1 and p2.
    Dim e1 As tVec3, e2 As tVec3
    Dim n1 As tVec3, n2 As tVec3

    e1 = pvWorld2EYE(W1)
    e2 = pvWorld2EYE(W2)
    If pvClipEYE(e1, e2) Then
        n1 = pvEYE2Norm(e1)
        n2 = pvEYE2Norm(e2)
        If pvClipNorm(n1, n2) Then
            P1 = pvNorm2Screen(n1)
            P2 = pvNorm2Screen(n2)
            LineToScreen = True: Exit Function
        End If
    End If

    P1 = PointToScreen(W1)
    P2 = PointToScreen(W2)
End Function
