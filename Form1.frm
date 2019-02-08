VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "3D Fishes  by reexre       (You can resize the Form)"
   ClientHeight    =   9225
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   9645
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   240
      Max             =   500
      Min             =   100
      TabIndex        =   4
      Top             =   3000
      Value           =   100
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkFollow 
      Caption         =   "Follow RND fish"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ResetCamera"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdR 
      Caption         =   "Re-Start"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.CheckBox chk3D 
      Caption         =   "3D View"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lZOOM 
      Caption         =   "ZOOM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents tmrClock As cTimer
Attribute tmrClock.VB_VarHelpID = -1

Private Pitch     As Double
Private Yaw       As Double

Private Const Deg2Rad As Double = 3.14159265358979 / 180#

Private Sub chk3D_Click()

    Do3D = chk3D.Value = vbChecked

    If Not (Do3D) Then Fish3Dto2D

    chkFollow.Visible = Do3D
    HScroll1.Visible = Do3D
    lZOOM.Visible = Do3D


End Sub



Private Sub chkFollow_Click()
    If chkFollow.Value = vbChecked Then
        CamTarget = Int(Rnd * NF) + 1
    Else
        CamTarget = 0
        Camera.cTo = Vec3(0, 0, 0)
        UpdateCamera
    End If


End Sub

Private Sub cmdR_Click()
    InitFishes NF

End Sub

Private Sub Command1_Click()
    InitCamera Vec3(0, 0, 1000), Vec3(0, 0, 0)

    Pitch = 0
    Yaw = 0    '90

    HScroll1.Value = 100

End Sub

Private Sub Form_DblClick()
    If Me.WindowState <> vbMaximized Then
        Me.WindowState = vbMaximized
    Else
        Me.WindowState = vbNormal
    End If
End Sub

Private Sub Form_Load()

    Randomize Timer
    InitFishes 250


    Scree.Center = Vec3(0, 0, 0)
    Scree.Size = Vec3(SideHalf, SideHalf, SideHalf)


    InitCamera Vec3(0, 0, 1000), Vec3(0, 0, 0)
    Command1_Click


    Set tmrClock = New_c.Timer(40, True)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim D         As Double
    Dim A         As Single

    Static x0, y0, dx, dy: dx = X - x0: dy = Y - y0

    Select Case Button
    Case 0
        x0 = X: y0 = Y
    Case 1
        If chk3D.Value Then    'don't modify pitch and yaw when we're not in 3D mode

            Pitch = Pitch - 0.25 * dy
            Yaw = (Yaw - 0.25 * dx)    ' Mod 360
            x0 = X: y0 = Y
            If Pitch > 90 Then Pitch = 90
            If Pitch < -90 Then Pitch = -90
            CameraSetRotation Yaw, Pitch

        End If

    Case 2    'zoom
        D = Sqr(vec3LEN2(vec3SUB(Camera.cFrom, Camera.cTo)))
        D = D - dy * 0.25
        If D < SideHalf * 1.7 Then D = SideHalf * 1.7

        With Camera
            .cFrom = vec3SUM(vec3MUL(Vec3Normalize(vec3SUB(.cFrom, .cTo)), D), .cTo)
        End With
        UpdateCamera

    End Select

End Sub

Private Sub Form_Resize()
    ScaleMode = vbPixels
    Set BBuf = Cairo.CreateSurface(ScaleWidth, ScaleHeight)
    ' RedrawOn BBuf.CreateContext
    ' Set Picture = BBuf.Picture    'finally set the updated content of the BackBuf-Surface as the new Form-Picture
End Sub

Private Sub HScroll1_Change()
    Camera.Zoom = HScroll1.Value * 0.01
    UpdateCamera
    lZOOM = "ZOOM: " & HScroll1.Value * 0.01
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub tmrClock_Timer()

'    If CamTarget Then
'        Camera.cTo = GetFishPos(CamTarget)
'        UpdateCamera
'    End If


    RedrawOn BBuf.CreateContext
    Set Picture = BBuf.Picture    'finally set the updated content of the BackBuf-Surface as the new Form-Picture



End Sub

Private Sub Form_Terminate()
    New_c.CleanupRichClientDll
End Sub
