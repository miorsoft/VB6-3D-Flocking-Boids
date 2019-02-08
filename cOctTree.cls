VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOCTTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Author; reexre
' OctTree for 3D collision detection

Private Type tCube
    x1            As Double
    y1            As Double
    z1            As Double
    x2            As Double
    y2            As Double
    z2            As Double
End Type

Private Boundary  As tCube

Private ChildNW   As cOCTTree
Private ChildNE   As cOCTTree
Private ChildSW   As cOCTTree
Private ChildSE   As cOCTTree

Private ChildNW2  As cOCTTree
Private ChildNE2  As cOCTTree
Private ChildSW2  As cOCTTree
Private ChildSE2  As cOCTTree


Private mCapacity As Long

Private mNP       As Long

Private pX()      As Double
Private pY()      As Double
Private pZ()      As Double
Private pIDX()    As Long
Private foundCount As Long
Private FoundUpperBound As Long

Private Const DefaultFoundUpperBound As Long = 1    '8

Public Divided    As Boolean


Private Sub SubDivide()
    Dim cnX       As Double
    Dim cnY       As Double
    Dim cnZ       As Double


    If ChildNW Is Nothing Then Set ChildNW = New cOCTTree
    If ChildNE Is Nothing Then Set ChildNE = New cOCTTree
    If ChildSW Is Nothing Then Set ChildSW = New cOCTTree
    If ChildSE Is Nothing Then Set ChildSE = New cOCTTree

    If ChildNW2 Is Nothing Then Set ChildNW2 = New cOCTTree
    If ChildNE2 Is Nothing Then Set ChildNE2 = New cOCTTree
    If ChildSW2 Is Nothing Then Set ChildSW2 = New cOCTTree
    If ChildSE2 Is Nothing Then Set ChildSE2 = New cOCTTree


    With Boundary
        cnX = (.x2 + .x1) * 0.5
        cnY = (.y2 + .y1) * 0.5
        cnZ = (.z2 + .z1) * 0.5

        ChildNW.Setup .x1, .y1, .z1, cnX, cnY, cnZ, mCapacity
        ChildNE.Setup cnX, .y1, .z1, .x2, cnY, cnZ, mCapacity
        ChildSW.Setup .x1, cnY, .z1, cnX, .y2, cnZ, mCapacity
        ChildSE.Setup cnX, cnY, .z1, .x2, .y2, cnZ, mCapacity

        ChildNW2.Setup .x1, .y1, cnZ, cnX, cnY, .z2, mCapacity
        ChildNE2.Setup cnX, .y1, cnZ, .x2, cnY, .z2, mCapacity
        ChildSW2.Setup .x1, cnY, cnZ, cnX, .y2, .z2, mCapacity
        ChildSE2.Setup cnX, cnY, cnZ, .x2, .y2, .z2, mCapacity

    End With

    Divided = True

End Sub


Friend Sub Setup(x1 As Double, y1 As Double, z1 As Double, _
                 x2 As Double, y2 As Double, z2 As Double, _
                 Capacity As Long)

    If Capacity Then
        With Boundary
            .x1 = x1
            .y1 = y1
            .z1 = z1
            .x2 = x2
            .y2 = y2
            .z2 = z2
        End With
        mCapacity = Capacity

        ReDim pX(mCapacity)
        ReDim pY(mCapacity)
        ReDim pZ(mCapacity)

        ReDim pIDX(mCapacity)
    End If

    Divided = False
    mNP = 0

End Sub



Friend Function InsertSinglePoint(pointX As Double, pointY As Double, pointZ As Double, _
                                  pointIDX As Long) As Boolean

    If Not (BoundaryContainPoint(pointX, pointY, pointZ)) Then Exit Function

    If mNP < mCapacity Then
        mNP = mNP + 1
        pX(mNP) = pointX
        pY(mNP) = pointY
        pZ(mNP) = pointZ
        pIDX(mNP) = pointIDX
        InsertSinglePoint = True

    Else


        If Not (Divided) Then SubDivide

        If ChildNW.InsertSinglePoint(pointX, pointY, pointZ, pointIDX) Then InsertSinglePoint = True: Exit Function
        If ChildNE.InsertSinglePoint(pointX, pointY, pointZ, pointIDX) Then InsertSinglePoint = True: Exit Function
        If ChildSW.InsertSinglePoint(pointX, pointY, pointZ, pointIDX) Then InsertSinglePoint = True: Exit Function
        If ChildSE.InsertSinglePoint(pointX, pointY, pointZ, pointIDX) Then InsertSinglePoint = True: Exit Function

        If ChildNW2.InsertSinglePoint(pointX, pointY, pointZ, pointIDX) Then InsertSinglePoint = True: Exit Function
        If ChildNE2.InsertSinglePoint(pointX, pointY, pointZ, pointIDX) Then InsertSinglePoint = True: Exit Function
        If ChildSW2.InsertSinglePoint(pointX, pointY, pointZ, pointIDX) Then InsertSinglePoint = True: Exit Function
        If ChildSE2.InsertSinglePoint(pointX, pointY, pointZ, pointIDX) Then InsertSinglePoint = True: Exit Function


    End If

End Function

Friend Sub InsertPoints(pointX() As Double, pointY() As Double, pointZ() As Double)
    Dim i         As Long
    Dim N         As Long
    Dim pntX      As Double
    Dim pntY      As Double
    Dim pntZ      As Double

    N = UBound(pointX)
    For i = 1 To N

        pntX = pointX(i)
        pntY = pointY(i)
        pntZ = pointZ(i)


        If (BoundaryContainPoint(pntX, pntY, pntZ)) Then
            If mNP < mCapacity Then
                mNP = mNP + 1
                pX(mNP) = pntX
                pY(mNP) = pntY
                pZ(mNP) = pntZ
                pIDX(mNP) = i

            Else
                If Not (Divided) Then SubDivide

                If Not (ChildNW.InsertSinglePoint(pntX, pntY, pntZ, i)) Then _
                   If Not (ChildNE.InsertSinglePoint(pntX, pntY, pntZ, i)) Then _
                   If Not (ChildSW.InsertSinglePoint(pntX, pntY, pntZ, i)) Then _
                   If Not (ChildSE.InsertSinglePoint(pntX, pntY, pntZ, i)) Then _
                   If Not (ChildNW2.InsertSinglePoint(pntX, pntY, pntZ, i)) Then _
                   If Not (ChildNE2.InsertSinglePoint(pntX, pntY, pntZ, i)) Then _
                   If Not (ChildSW2.InsertSinglePoint(pntX, pntY, pntZ, i)) Then _
                   ChildSE2.InsertSinglePoint pntX, pntY, pntZ, i

            End If
        End If
    Next

End Sub

Private Function BoundaryContainPoint(X As Double, Y As Double, Z As Double) As Boolean

    With Boundary
        If X > .x2 Then Exit Function
        If X < .x1 Then Exit Function
        If Y > .y2 Then Exit Function
        If Y < .y1 Then Exit Function
        If Z > .z2 Then Exit Function
        If Z < .z1 Then Exit Function

    End With
    BoundaryContainPoint = True

End Function


'''Friend Sub DRAW(ShowQuads As Long)
'''    Dim I   As Long
'''    Dim x   As Double
'''    Dim y   As Double
'''
'''    If mNP = 0 Then Exit Sub
'''
'''    If ShowQuads Then
'''        '    'DrawCross
'''        With Boundary
'''            x = (.x2 + .x1) * 0.5
'''            y = (.y2 + .y1) * 0.5
'''            vbDrawCC.DrawLine x, .y1, x, .y2, , 1, vbWhite, 0.2
'''            vbDrawCC.DrawLine .x1, y, .x2, y, , 1, vbWhite, 0.2
'''        End With
'''    End If
'''
'''    'DrawPoints
'''    vbDrawCC.SetSourceColor vbYellow, 0.8
'''    For I = 1 To mNP
'''        vbDrawCC.Arc pX(I), pY(I), 2
'''        vbDrawCC.stroke
'''    Next
'''
'''
'''    If Divided Then
'''        ChildNW.DRAW ShowQuads
'''        ChildNE.DRAW ShowQuads
'''        ChildSW.DRAW ShowQuads
'''        ChildSE.DRAW ShowQuads
'''    End If
'''
'''
'''End Sub





Private Function DoCubesIntersect(C1minX As Double, C1minY As Double, C1minZ As Double, _
                                  C1MaxX As Double, C1MaxY As Double, C1MaxZ As Double, _
                                  C2minX As Double, C2minY As Double, C2minZ As Double, _
                                  C2MaxX As Double, C2MaxY As Double, C2MaxZ As Double) As Boolean
    If C1MaxX < C2minX Then Exit Function
    If C1MaxY < C2minY Then Exit Function
    If C1minX > C2MaxX Then Exit Function
    If C1minY > C2MaxY Then Exit Function

    If C1MaxZ < C2minZ Then Exit Function
    If C1minZ > C2MaxZ Then Exit Function

    DoCubesIntersect = True

End Function

Friend Sub QueryCube(rX1 As Double, _
                     ry1 As Double, _
                     rZ1 As Double, _
                     rX2 As Double, _
                     ry2 As Double, _
                     rZ2 As Double, _
                     rpX() As Double, rpY() As Double, rpZ() As Double, _
                     rpIDX() As Long)


    FoundUpperBound = DefaultFoundUpperBound
    ReDim rpX(FoundUpperBound)
    ReDim rpY(FoundUpperBound)
    ReDim rpZ(FoundUpperBound)
    ReDim rpIDX(FoundUpperBound)

    foundCount = 0

    pvQueryCube rX1, ry1, rZ1, _
                rX2, ry2, rZ2, _
                rpX(), rpY(), rpZ(), _
                rpIDX(), foundCount, FoundUpperBound

    ReDim Preserve rpX(foundCount)
    ReDim Preserve rpY(foundCount)
    ReDim Preserve rpZ(foundCount)
    ReDim Preserve rpIDX(foundCount)

End Sub


Friend Sub pvQueryCube(rX1 As Double, _
                       ry1 As Double, _
                       rZ1 As Double, _
                       rX2 As Double, _
                       ry2 As Double, _
                       rZ2 As Double, _
                       rpX() As Double, rpY() As Double, rpZ() As Double, _
                       rpIDX() As Long, foundCount As Long, FoundUpperBound As Long)


    Dim i         As Long

    If Not (DoCubesIntersect(rX1, ry1, rZ1, _
                             rX2, ry2, rZ2, _
                             Boundary.x1, Boundary.y1, Boundary.z1, _
                             Boundary.x2, Boundary.y2, Boundary.z2)) Then Exit Sub


    For i = 1 To mNP    'Point in Range ?
        If pX(i) >= rX1 Then
            If pX(i) < rX2 Then

                If pY(i) >= ry1 Then
                    If pY(i) < ry2 Then

                        If pZ(i) >= rZ1 Then
                            If pZ(i) < rZ2 Then

                                foundCount = foundCount + 1
                                If foundCount > FoundUpperBound Then
                                    FoundUpperBound = foundCount * 2
                                    ReDim Preserve rpX(FoundUpperBound)
                                    ReDim Preserve rpY(FoundUpperBound)
                                    ReDim Preserve rpZ(FoundUpperBound)
                                    ReDim Preserve rpIDX(FoundUpperBound)
                                End If

                                rpX(foundCount) = pX(i)
                                rpY(foundCount) = pY(i)
                                rpZ(foundCount) = pZ(i)
                                rpIDX(foundCount) = pIDX(i)
                            End If
                        End If

                    End If
                End If

            End If
        End If

    Next

    If Divided Then
        ChildNW.pvQueryCube rX1, ry1, rZ1, rX2, ry2, rZ2, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound
        ChildNE.pvQueryCube rX1, ry1, rZ1, rX2, ry2, rZ2, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound
        ChildSW.pvQueryCube rX1, ry1, rZ1, rX2, ry2, rZ2, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound
        ChildSE.pvQueryCube rX1, ry1, rZ1, rX2, ry2, rZ2, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound

        ChildNW2.pvQueryCube rX1, ry1, rZ1, rX2, ry2, rZ2, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound
        ChildNE2.pvQueryCube rX1, ry1, rZ1, rX2, ry2, rZ2, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound
        ChildSW2.pvQueryCube rX1, ry1, rZ1, rX2, ry2, rZ2, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound
        ChildSE2.pvQueryCube rX1, ry1, rZ1, rX2, ry2, rZ2, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound
    End If

End Sub








Friend Sub QuerySphere(cx As Double, _
                       cy As Double, _
                       cz As Double, _
                       R As Double, _
                       rpX() As Double, rpY() As Double, rpZ() As Double, _
                       rpIDX() As Long)

    FoundUpperBound = DefaultFoundUpperBound
    ReDim rpX(FoundUpperBound)
    ReDim rpY(FoundUpperBound)
    ReDim rpZ(FoundUpperBound)
    ReDim rpIDX(FoundUpperBound)

    foundCount = 0

    pvQuerySphere cx, cy, cz, R, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound

    ReDim Preserve rpX(foundCount)
    ReDim Preserve rpY(foundCount)
    ReDim Preserve rpZ(foundCount)
    ReDim Preserve rpIDX(foundCount)

End Sub

Friend Sub pvQuerySphere(cx As Double, _
                         cy As Double, _
                         cz As Double, _
                         R As Double, _
                         rpX() As Double, rpY() As Double, rpZ() As Double, _
                         rpIDX() As Long, foundCount As Long, FoundUpperBound As Long)


    Dim i         As Long
    Dim rX1       As Double: rX1 = cx - R
    Dim ry1       As Double: ry1 = cy - R
    Dim rX2       As Double: rX2 = cx + R
    Dim ry2       As Double: ry2 = cy + R

    Dim rZ1       As Double: rZ1 = cz - R
    Dim rZ2       As Double: rZ2 = cz + R

    Dim dx        As Double
    Dim dy        As Double
    Dim Dz        As Double

    Dim rSQ       As Double
    rSQ = R * R

    If Not (DoCubesIntersect(rX1, ry1, rZ1, _
                             rX2, ry2, rZ2, _
                             Boundary.x1, Boundary.y1, Boundary.z1, _
                             Boundary.x2, Boundary.y2, Boundary.z2)) Then Exit Sub

    For i = 1 To mNP
        dx = pX(i) - cx
        dy = pY(i) - cy
        Dz = pZ(i) - cz
        If (dx * dx + dy * dy + Dz * Dz) < rSQ Then
            foundCount = foundCount + 1
            If foundCount > FoundUpperBound Then
                FoundUpperBound = foundCount * 2
                ReDim Preserve rpX(FoundUpperBound)
                ReDim Preserve rpY(FoundUpperBound)
                ReDim Preserve rpZ(FoundUpperBound)
                ReDim Preserve rpIDX(FoundUpperBound)
            End If

            rpX(foundCount) = pX(i)
            rpY(foundCount) = pY(i)
            rpZ(foundCount) = pZ(i)
            rpIDX(foundCount) = pIDX(i)

        End If
    Next

    If Divided Then
        ChildNW.pvQuerySphere cx, cy, cz, R, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound
        ChildNE.pvQuerySphere cx, cy, cz, R, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound
        ChildSW.pvQuerySphere cx, cy, cz, R, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound
        ChildSE.pvQuerySphere cx, cy, cz, R, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound
        ChildNW2.pvQuerySphere cx, cy, cz, R, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound
        ChildNE2.pvQuerySphere cx, cy, cz, R, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound
        ChildSW2.pvQuerySphere cx, cy, cz, R, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound
        ChildSE2.pvQuerySphere cx, cy, cz, R, rpX(), rpY(), rpZ(), rpIDX(), foundCount, FoundUpperBound

    End If

End Sub


















'********************************************************************
'Query for 2D collision:

''Q.Setup 0, 0, MaxW * 1, maxH * 1, 30
'''For I = 1 To mNP
'''    Q.InsertSinglePoint x(I), y(I), I
'''Next
''Q.inserpoints x, y
''
''
''
''For I = 1 To mNP
''    Q.QueryCube x(I) - R * 2, y(I) - R * 2, _
 ''            x(I) + R * 2, y(I) + R * 2, rX(), rY(), rIDX()
''
''    For J = 1 To UBound(rX)
''        If I < rIDX(J) Then
''            dx = x(I) - rX(J)
''            dy = y(I) - rY(J)
''            If dx * dx + dy * dy < diam2 Then  'Diam*Diam   (2*R)^2
''                'COLLISION HAPPENED
''            End If
''        End If
''
''    Next
''Next


