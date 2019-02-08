Attribute VB_Name = "m3DtoScreen"
Option Explicit


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
Private Const DTOR As Double = 0.01745329252    'Degrees to Radians

Private CamNormrightDIR As tVec3
Private CamNormDIR As tVec3
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

        .ANGh = 40
        .ANGv = 40 * Scree.Size.Y / Scree.Size.X
        .Projection = PERSPECTIVE

        .Zoom = 1

        INVtanthetaH = 1 / Tan(.ANGh * DTOR * 0.5)
        INVtanthetaV = 1 / Tan(.ANGv * DTOR * 0.5)
    End With

    UpdateCamera

End Sub
Public Function UpdateCamera()

    CamNormDIR.X = Camera.cTo.X - Camera.cFrom.X
    CamNormDIR.Y = Camera.cTo.Y - Camera.cFrom.Y
    CamNormDIR.Z = Camera.cTo.Z - Camera.cFrom.Z

    CamNormDIR = Vec3Normalize(CamNormDIR)

    CamNormrightDIR = Vec3Cross(Camera.cUp, CamNormDIR)
    CamNormrightDIR = Vec3Normalize(CamNormrightDIR)

    CamNormUPDir = Vec3Cross(CamNormDIR, CamNormrightDIR)

    '   /* Calculate camera aperture statics, note: angles in degrees */
    '   tanthetah = tan(camera.angleh * DTOR / 2);
    '   tanthetav = tan(camera.anglev * DTOR / 2);

    '    INVtanthetaH = 1 / Tan(Camera.ANGh * DTOR * 0.5)
    '    INVtanthetaV = 1 / Tan(Camera.ANGv * DTOR * 0.5)

End Function


' Take a point in world coordinates and transform it to
'   a point in the eye coordinate system.

'void Trans_World2Eye(W, e, CAMERA)
'XYZ w;
'XYZ *e;
'CAMERA camera;
'{
'   /* Translate world so that the camera is at the origin */
'   w.x -= camera.from.x;
'   w.y -= camera.from.y;
'   w.z -= camera.from.z;
'
'   /* Convert to eye coordinates using basis vectors */
'   e->x = w.x * CamNormRightDir.x + w.y * CamNormRightDir.y + w.z * CamNormRightDir.z;
'   e->y = w.x * CamNormDIR.x + w.y * CamNormDIR.y + w.z * CamNormDIR.z;
'   e->z = w.x * CamNormUPDir.x + w.y * CamNormUPDir.y + w.z * CamNormUPDir.z;
'}
Public Function World2EYE(inW As tVec3) As tVec3
    Dim W         As tVec3

    With W
        .X = inW.X - Camera.cFrom.X
        .Y = inW.Y - Camera.cFrom.Y
        .Z = inW.Z - Camera.cFrom.Z

        World2EYE.X = .X * CamNormrightDIR.X + .Y * CamNormrightDIR.Y + .Z * CamNormrightDIR.Z
        World2EYE.Y = .X * CamNormDIR.X + .Y * CamNormDIR.Y + .Z * CamNormDIR.Z
        World2EYE.Z = .X * CamNormUPDir.X + .Y * CamNormUPDir.Y + .Z * CamNormUPDir.Z
    End With



End Function



Private Function pvClipEYE(E1 As tVec3, E2 As tVec3) As Boolean


    Dim Mu        As Double
    '   /* Is the vector totally in NearPlane of the NearPlane cutting plane ? */
    If (E1.Y <= Camera.NearPlane And E2.Y <= Camera.NearPlane) Then pvClipEYE = False: Exit Function

    '   /* Is the vector totally behind the FarPlane cutting plane ? */
    If (E1.Y >= Camera.FarPlane And E2.Y >= Camera.FarPlane) Then pvClipEYE = False: Exit Function

    '   /* Is the vector partly in NearPlane of the NearPlane cutting plane ? */
    If ((E1.Y < Camera.NearPlane And E2.Y > Camera.NearPlane) Or _
        (E1.Y > Camera.NearPlane And E2.Y < Camera.NearPlane)) Then
        Mu = (Camera.NearPlane - E1.Y) / (E2.Y - E1.Y)
        If (E1.Y < Camera.NearPlane) Then
            E1.X = E1.X + Mu * (E2.X - E1.X)
            E1.Z = E1.Z + Mu * (E2.Z - E1.Z)
            E1.Y = Camera.NearPlane
        Else
            E2.X = E1.X + Mu * (E2.X - E1.X)
            E2.Z = E1.Z + Mu * (E2.Z - E1.Z)
            E2.Y = Camera.NearPlane
        End If
    End If

    '   /* Is the vector partly behind the FarPlane cutting plane ? */
    If ((E1.Y < Camera.FarPlane And E2.Y > Camera.FarPlane) Or _
        (E1.Y > Camera.FarPlane And E2.Y < Camera.FarPlane)) Then
        Mu = (Camera.FarPlane - E1.Y) / (E2.Y - E1.Y)
        If (E1.Y < Camera.FarPlane) Then
            E2.X = E1.X + Mu * (E2.X - E1.X)
            E2.Z = E1.Z + Mu * (E2.Z - E1.Z)
            E2.Y = Camera.FarPlane
        Else
            E1.X = E1.X + Mu * (E2.X - E1.X)
            E1.Z = E1.Z + Mu * (E2.Z - E1.Z)
            E1.Y = Camera.FarPlane
        End If
    End If

    pvClipEYE = True

End Function


Private Function pvEYE2Norm(E As tVec3) As tVec3
    Dim D         As Double

    If Camera.Projection = PERSPECTIVE Then
        D = Camera.Zoom / E.Y
        pvEYE2Norm.X = D * E.X * INVtanthetaH
        pvEYE2Norm.Y = E.Y
        pvEYE2Norm.Z = D * E.Z * INVtanthetaV
    Else
        pvEYE2Norm.X = Camera.Zoom * E.X * INVtanthetaH
        pvEYE2Norm.Y = E.Y
        pvEYE2Norm.Z = Camera.Zoom * E.Z * INVtanthetaV
    End If


End Function






''''/*
''''   Clip a line segment to the normalised coordinate +- square.
''''   The y component is not touched.
''''*/
'''Public Function ClipNorm(ByRef n1 As tVec3, n2 As tVec3) As Boolean
'''    Dim Mu  As Double
'''
'''    '   /* Is the line segment totally right of x = 1 ? */
'''    If (n1.x >= 1 And n2.x >= 1) Then ClipNorm = False: Exit Function
'''
'''    '   /* Is the line segment totally left of x = -1 ? */
'''    If (n1.x <= -1 And n2.x <= -1) Then ClipNorm = False: Exit Function
'''
'''    '   /* Does the vector cross x = 1 ? */
'''    If ((n1.x > 1 And n2.x < 1) Or (n1.x < 1 And n2.x > 1)) Then
'''        Mu = (1 - n1.x) / (n2.x - n1.x)
'''        If (n1.x < 1) Then
'''            n2.Z = n1.Z + Mu * (n2.Z - n1.Z)
'''            n2.x = 1
'''        Else
'''            n1.Z = n1.Z + Mu * (n2.Z - n1.Z)
'''            n1.x = 1
'''        End If
'''    End If
'''
'''    '   /* Does the vector cross x = -1 ? */
'''    If ((n1.x < -1 And n2.x > -1) Or (n1.x > -1 And n2.x < -1)) Then
'''        Mu = (-1 - n1.x) / (n2.x - n1.x)
'''        If (n1.x > -1) Then
'''            n2.Z = n1.Z + Mu * (n2.Z - n1.Z)
'''            n2.x = -1
'''        Else
'''            n1.Z = n1.Z + Mu * (n2.Z - n1.Z)
'''            n1.x = -1
'''        End If
'''    End If
'''
'''    '   /* Is the line segment totally above z = 1 ? */
'''    If (n1.Z >= 1 And n2.Z >= 1) Then ClipNorm = False: Exit Function
'''
'''    '   /* Is the line segment totally below z = -1 ? */
'''    If (n1.Z <= -1 And n2.Z <= -1) Then ClipNorm = False: Exit Function
'''
'''    '   /* Does the vector cross z = 1 ? */
'''    If ((n1.Z > 1 And n2.Z < 1) Or (n1.Z < 1 And n2.Z > 1)) Then
'''        Mu = (1 - n1.Z) / (n2.Z - n1.Z)
'''        If (n1.Z < 1) Then
'''            n2.x = n1.x + Mu * (n2.x - n1.x)
'''            n2.Z = 1
'''        Else
'''            n1.x = n1.x + Mu * (n2.x - n1.x)
'''            n1.Z = 1
'''        End If
'''    End If
'''
'''    '   /* Does the vector cross z = -1 ? */
'''    If ((n1.Z < -1 And n2.Z > -1) Or (n1.Z > -1 And n2.Z < -1)) Then
'''        Mu = (-1 - n1.Z) / (n2.Z - n1.Z)
'''        If (n1.Z > -1) Then
'''            n2.x = n1.x + Mu * (n2.x - n1.x)
'''            n2.Z = -1
'''        Else
'''            n1.x = n1.x + Mu * (n2.x - n1.x)
'''            n1.Z = -1
'''        End If
'''    End If
'''
'''
'''    ClipNorm = True
'''
'''End Function



Private Function pvNorm2Screen(norm As tVec3) As tVec3

'pvNorm2Screen.x = Scree.Center.x + Scree.Size.x * norm.x * 0.5
    pvNorm2Screen.X = Scree.Center.X - Scree.Size.X * norm.X * 0.5    'Not sure about X sign
    pvNorm2Screen.Y = Scree.Center.Y - Scree.Size.Y * norm.Z * 0.5

End Function


'/*
'   Transform a point from world to screen coordinates. Return TRUE
'   if the point is visible, the point in screen coordinates is p.
'   Assumes Trans_Initialise() has been called
'*/
'int Trans_Point(w,p,screen,camera)
'XYZ w;
'Point *p;
'SCREEN screen;
'CAMERA camera;
'{
'   XYZ e,n;
'
'   Trans_World2Eye(w,&e,camera);
'   if (e.y >= camera.NearPlane && e.y <= camera.FarPlane) {
'      Trans_pvEYE2Norm(e,&n,camera);
'      if (n.x >= -1 && n.x <= 1 && n.z >= -1 && n.z <= 1) {
'         Trans_pvNorm2Screen(n,p,screen);
'         return(TRUE);
'      }
'   }
'   return(FALSE);
'}





''''Public Function LineToScreen(w1 As tVec3, w2 As tVec3, RetP1 As tVec2, RetP2 As tVec2) As Boolean
''''
''''    Dim E1  As tVec3
''''    Dim E2  As tVec3
''''    Dim n1  As tVec3
''''    Dim n2  As tVec3
''''
''''    E1 = World2EYE(w1)
''''    E2 = World2EYE(w2)
''''    If pvClipEYE(E1, E2) Then
''''        n1 = pvEYE2Norm(E1)
''''        n2 = pvEYE2Norm(E2)
''''        If ClipNorm(n1, n2) Then
''''            RetP1 = pvNorm2Screen(n1)
''''            RetP2 = pvNorm2Screen(n2)
''''            LineToScreen = True
''''        Else
''''            LineToScreen = False
''''        End If
''''    Else
''''        LineToScreen = False
''''    End If
''''End Function

Public Function DistFromCamera2(V As tVec3) As Double
    Dim dx        As Double
    Dim dy        As Double
    Dim Dz        As Double
    dx = V.X - Camera.cFrom.X
    dy = V.Y - Camera.cFrom.Y
    Dz = V.Z - Camera.cFrom.Z
    DistFromCamera2 = dx * dx + dy * dy + Dz * Dz
End Function




Public Function PointToScreenOLD(W As tVec3) As tVec3

    Dim E         As tVec3
    Dim N         As tVec3

    E = World2EYE(W)
    'If (E.y >= camera.NearPlane And E.y <= camera.FarPlane) Then
    N = pvEYE2Norm(E)
    '        Stop
    '''    If (N.X >= -2 And N.X <= 2 And N.Z >= -2 And N.Z <= 2) Then    '[1]

    PointToScreenOLD = pvNorm2Screen(N)

End Function


Public Function PointToScreen(W As tVec3) As tVec3
' Put all just in 1 Call

    Dim PO        As tVec3
    Dim EYE       As tVec3
    Dim eyeN      As tVec3
    Dim D         As Double


    '''   Take a point in world coordinates and transform it to
    '''   a point in the eye coordinate system.
    PO = vec3SUB(W, Camera.cFrom)
    EYE.X = vec3DOT(PO, CamNormrightDIR)    'Projection
    EYE.Y = vec3DOT(PO, CamNormDIR)
    EYE.Z = vec3DOT(PO, CamNormUPDir)


    '    If vec3DOT(PO, CamNormDIR) < 0 Then EYE.Y = 2 / EYE.Y
    '--------------


    ''   Take a vector in eye coordinates and transform it into
    ''   normalised coordinates for a perspective view. No normalisation
    ''   is performed for an orthographic projection. Note that although
    ''   the y component of the normalised vector is copied from the eye
    ''   coordinate system, it is generally no longer needed. It can
    ''   however still be used externally for vector sorting.
    ' Eye to NORM
    If Camera.Projection = PERSPECTIVE Then
        D = Camera.Zoom / EYE.Y
        eyeN.X = D * EYE.X * INVtanthetaH
        eyeN.Y = EYE.Y
        eyeN.Z = D * EYE.Z * INVtanthetaV
    Else
        eyeN.X = Camera.Zoom * EYE.X * INVtanthetaH
        eyeN.Y = EYE.Y
        eyeN.Z = Camera.Zoom * EYE.Z * INVtanthetaV
    End If


    PointToScreen.X = Scree.Center.X - Scree.Size.X * eyeN.X * 0.5    'Not sure about X sign
    PointToScreen.Y = Scree.Center.Y - Scree.Size.Y * eyeN.Z * 0.5

    PointToScreen.Z = eyeN.Z

End Function





