Attribute VB_Name = "mBehaviors"
Option Explicit

Public Type tBasicBehavior
    Cohesion      As Double
    Separation    As Double
    InvSeparation As Double
    Alignment     As Double
End Type

Public Type tBehavior
    Strength      As tBasicBehavior
    Distances     As tBasicBehavior
    R             As Double
    G             As Double
    B             As Double
End Type

Public BehaviorTable() As tBehavior


Public NofBehaviors As Long

Public MaxInteractDist As Double
Public MaxInteractDist2 As Double


Public Sub InitBehaviors()
    Dim i         As Long
    Dim J         As Long

    NofBehaviors = 3

    ' They are how do MyType behave respect to Other Type, so:
    ReDim BehaviorTable(NofBehaviors, NofBehaviors)


    For i = 1 To NofBehaviors
        For J = 1 To NofBehaviors
            With BehaviorTable(i, J)

                '                'V1
                '                .Distances.Alignment = 110
                '                .Distances.Cohesion = 80
                '                .Distances.Separation = 50
                '                .Strength.Alignment = 0.08
                '                .Strength.Cohesion = 0.1
                '                .Strength.Separation = 1
                .Distances.Alignment = 80
                .Distances.Cohesion = 100
                .Distances.Separation = 55
                .Strength.Alignment = 0.04
                .Strength.Cohesion = 0.025
                .Strength.Separation = 0.5





                'To make Behavior I random respect to J
                '                                .Strength.Alignment = .Strength.Alignment * (0 + Rnd)
                '                                .Strength.Cohesion = .Strength.Cohesion * (-1 + Rnd * 2)
                '                                .Strength.Separation = .Strength.Separation * (0.5 + Rnd)


                .Distances.Alignment = .Distances.Alignment * .Distances.Alignment
                .Distances.Cohesion = .Distances.Cohesion * .Distances.Cohesion
                .Distances.Separation = .Distances.Separation * .Distances.Separation
                .Distances.InvSeparation = 1 / (.Distances.Separation)

                If .Distances.Alignment > MaxInteractDist2 Then MaxInteractDist2 = .Distances.Alignment
                If .Distances.Cohesion > MaxInteractDist2 Then MaxInteractDist2 = .Distances.Cohesion
                If .Distances.Separation > MaxInteractDist2 Then MaxInteractDist2 = .Distances.Separation

                Do
                    .R = Rnd: .G = Rnd: .B = Rnd
                Loop While .R * 3 + .G * 6 + .B * 3 < 4

            End With
        Next
    Next

    For i = 1 To NofBehaviors
        For J = 1 To NofBehaviors
            If i <> J Then
                BehaviorTable(i, J).Strength.Alignment = 0
                BehaviorTable(i, J).Strength.Separation = 0
                BehaviorTable(i, J).Strength.Cohesion = -50000 * BehaviorTable(i, J).Strength.Cohesion
            End If
        Next
    Next

    For i = 1 To NofBehaviors
        For J = 1 To NofBehaviors
            If (i = 2) Then
                If (J <> 2) Then  '' Make Group 2 be PREDATOR like with everyone else
                    BehaviorTable(i, J).Distances.Alignment = 10 * 10    '!! REMEBMER to POW 2
                    BehaviorTable(i, J).Distances.Cohesion = 65 * 65
                    BehaviorTable(i, J).Distances.Separation = 10 * 10
                    BehaviorTable(i, J).Strength.Alignment = 0
                    BehaviorTable(i, J).Strength.Cohesion = 0.5
                    BehaviorTable(i, J).Strength.Separation = 0

                Else
                    BehaviorTable(i, J).Distances.Alignment = 50 * 50
                    BehaviorTable(i, J).Distances.Cohesion = 100 * 100
                    BehaviorTable(i, J).Distances.Separation = 55 * 55
                    BehaviorTable(i, J).Strength.Alignment = 0.001
                    BehaviorTable(i, J).Strength.Cohesion = 0.33
                    BehaviorTable(i, J).Strength.Separation = 0.55
                End If
            End If

        Next
    Next

    MaxInteractDist = Sqr(MaxInteractDist2)

End Sub
