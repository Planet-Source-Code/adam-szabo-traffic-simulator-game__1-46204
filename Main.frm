VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traffic Simulator"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   579
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   726
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Main 
      Interval        =   50
      Left            =   7440
      Top             =   1200
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Up and Down Traffic"
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   6480
      TabIndex        =   4
      Top             =   5160
      Width           =   1815
      Begin VB.Timer tmrLightU 
         Enabled         =   0   'False
         Interval        =   2500
         Left            =   240
         Top             =   240
      End
      Begin VB.OptionButton optLightU 
         Caption         =   "Auto"
         Height          =   330
         Index           =   2
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optLightU 
         Caption         =   "<"
         Height          =   255
         Index           =   1
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1400
         Width           =   615
      End
      Begin VB.OptionButton optLightU 
         Caption         =   "<"
         Height          =   255
         Index           =   0
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   870
         Width           =   615
      End
      Begin VB.Shape GreenU 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   240
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   495
      End
      Begin VB.Shape RedU 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   240
         Shape           =   3  'Circle
         Top             =   840
         Width           =   495
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         Height          =   1095
         Left            =   240
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left and Right Traffic"
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   2400
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
      Begin VB.Timer tmrLight 
         Interval        =   2500
         Left            =   240
         Top             =   240
      End
      Begin VB.OptionButton optLight 
         Caption         =   "<"
         Height          =   255
         Index           =   0
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   870
         Width           =   615
      End
      Begin VB.OptionButton optLight 
         Caption         =   "<"
         Height          =   255
         Index           =   1
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1400
         Width           =   615
      End
      Begin VB.OptionButton optLight 
         Caption         =   "Auto"
         Height          =   330
         Index           =   2
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.Shape Red 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   240
         Shape           =   3  'Circle
         Top             =   840
         Width           =   495
      End
      Begin VB.Shape Green 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   240
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   495
      End
      Begin VB.Shape Light 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         Height          =   1095
         Left            =   240
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Image dCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   5
      Left            =   4680
      Picture         =   "Main.frx":0000
      Top             =   480
      Width           =   480
   End
   Begin VB.Image dCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   4680
      Picture         =   "Main.frx":0C42
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image dCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Left            =   4680
      Picture         =   "Main.frx":1884
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image dCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   2
      Left            =   4680
      Picture         =   "Main.frx":24C6
      Top             =   5400
      Width           =   480
   End
   Begin VB.Image dCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   4680
      Picture         =   "Main.frx":3108
      Top             =   6720
      Width           =   480
   End
   Begin VB.Image dCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   4680
      Picture         =   "Main.frx":3D4A
      Top             =   7440
      Width           =   480
   End
   Begin VB.Image uCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   5520
      Picture         =   "Main.frx":498C
      Top             =   7800
      Width           =   480
   End
   Begin VB.Image uCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Left            =   5520
      Picture         =   "Main.frx":55CE
      Top             =   6240
      Width           =   480
   End
   Begin VB.Image uCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   2
      Left            =   5520
      Picture         =   "Main.frx":6210
      Top             =   5520
      Width           =   480
   End
   Begin VB.Image uCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   5520
      Picture         =   "Main.frx":6E52
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image uCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   5520
      Picture         =   "Main.frx":7A94
      Top             =   600
      Width           =   480
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   2  'Dash
      Index           =   3
      X1              =   355
      X2              =   355
      Y1              =   344
      Y2              =   560
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   2  'Dash
      Index           =   2
      X1              =   355
      X2              =   355
      Y1              =   16
      Y2              =   232
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   296
      X2              =   416
      Y1              =   336
      Y2              =   336
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   296
      X2              =   416
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Image rCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   6
      Left            =   240
      Picture         =   "Main.frx":86D6
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image rCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   5
      Left            =   840
      Picture         =   "Main.frx":9318
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image rCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   2040
      Picture         =   "Main.frx":9F5A
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image rCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Left            =   3480
      Picture         =   "Main.frx":AB9C
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image rCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   2
      Left            =   6480
      Picture         =   "Main.frx":B7DE
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image rCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   7680
      Picture         =   "Main.frx":C420
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image rCar 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   9240
      Picture         =   "Main.frx":D062
      Top             =   4440
      Width           =   480
   End
   Begin VB.Image Car 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   6
      Left            =   10080
      Picture         =   "Main.frx":DCA4
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image Car 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   5
      Left            =   9360
      Picture         =   "Main.frx":E8E6
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image Car 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Left            =   6720
      Picture         =   "Main.frx":F528
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image Car 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   7800
      Picture         =   "Main.frx":1016A
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image Car 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   2
      Left            =   3840
      Picture         =   "Main.frx":10DAC
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image Car 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   2280
      Picture         =   "Main.frx":119EE
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image Car 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   720
      Picture         =   "Main.frx":12630
      Top             =   3720
      Width           =   480
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   2  'Dash
      Index           =   1
      X1              =   8
      X2              =   280
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   2  'Dash
      Index           =   0
      X1              =   416
      X2              =   712
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   296
      X2              =   296
      Y1              =   240
      Y2              =   336
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   416
      X2              =   416
      Y1              =   240
      Y2              =   336
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1455
      Left            =   120
      Top             =   3600
      Width           =   10575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   8415
      Left            =   4440
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
lStage = 2
lStageU = 1
With rdLine
    .Left = 417
    .Right = 418
    .Top = 240
    .Bottom = 336
End With
With rdLine1
    .Left = 296
    .Right = 297
    .Top = 240
    .Bottom = 336
End With
With rdLine2
    .Left = 296
    .Right = 416
    .Top = 238
    .Bottom = 239
End With
With rdLine3
    .Left = 296
    .Right = 416
    .Top = 339
    .Bottom = 340
End With
End Sub

Private Sub Main_Timer()
For x = 0 To nMaxCarUp
    With aiCarU(x)
        .Left = uCar(x).Left
        .Top = uCar(x).Top
        .Right = .Left + 32
        .Bottom = .Top + 32
    End With
Next x
For x = 0 To nMaxCarUp
    If x = 0 Then
        If IntersectRect(IntersectArea, aiCarU(x), aiCarU(nMaxCarUp)) Then
            aiMoveU(x) = False
        Else
            aiMoveU(x) = True
        End If
    Else
        If IntersectRect(IntersectArea, aiCarU(x), aiCarU(x - 1)) Then
            aiMoveU(x) = False
        Else
            aiMoveU(x) = True
        End If
    End If
Next x
For x = 0 To nMaxCarUp
    With aiCarU(x)
        .Left = uCar(x).Left
        .Top = uCar(x).Top
        .Right = .Left + 32
        .Bottom = .Top + 10
    End With
Next x
For x = 0 To nMaxCarUp
    If aiMoveU(x) = True Then
        If IntersectRect(IntersectArea, aiCarU(x), rdLine3) Then
            If lStageU > 1 Then 'Green
                'Check for collision with other traffic going right
                'by simulating the vehicle has moved forward and checking
                'a large area around the vehicle to make sure it has enough
                'room to advance without collision
                With aiCarU(x)
                    .Left = uCar(x).Left - 62
                    .Top = uCar(x).Top - 15
                    .Right = .Left + 100
                    .Bottom = .Top + 32
                End With
                For y = 0 To nMaxCarRight
                    If IntersectRect(IntersectArea, aiCarU(x), aiCarR(y)) Then
                        'if it would collide with a car going in that direction
                        GoTo dSkip
                    End If
                Next y
                uCar(x).Move uCar(x).Left, uCar(x).Top - 5
            End If
        Else
            uCar(x).Move uCar(x).Left, uCar(x).Top - 5
        End If
    End If
dSkip:
    If uCar(x).Top <= 16 Then
        uCar(x).Top = 528
    End If
Next x

'#############################################################################
For x = 0 To nMaxCarLeft
    With aiCar(x)
        .Left = Car(x).Left
        .Top = Car(x).Top
        .Right = .Left + 32
        .Bottom = .Top + 32
    End With
Next x
For x = 0 To nMaxCarLeft
    If x = 0 Then
        If IntersectRect(IntersectArea, aiCar(x), aiCar(nMaxCarLeft)) Then
            aiMove(x) = False
        Else
            aiMove(x) = True
        End If
    Else
        If IntersectRect(IntersectArea, aiCar(x), aiCar(x - 1)) Then
            aiMove(x) = False
        Else
            aiMove(x) = True
        End If
    End If
Next x
For x = 0 To nMaxCarLeft
    With aiCar(x)
        .Left = Car(x).Left
        .Top = Car(x).Top
        .Right = .Left + 10
        .Bottom = .Top + 32
    End With
Next x
For x = 0 To nMaxCarLeft
    If aiMove(x) = True Then
        If IntersectRect(IntersectArea, aiCar(x), rdLine) Then
            If lStage > 1 Then 'Green
                'Check for collision with other traffic going up
                'by simulating the vehicle has moved left and checking
                'a large area around the vehicle to make sure it has enough
                'room to advance without collision
                With aiCar(x)
                    .Left = Car(x).Left - 25
                    .Top = Car(x).Top - 25
                    .Right = .Left + 32
                    .Bottom = .Top + 95
                End With
                For y = 0 To nMaxCarUp
                    If IntersectRect(IntersectArea, aiCar(x), aiCarU(y)) Then
                        'if it would collide with a car going in that direction
                        GoTo lSkip
                    End If
                Next y
                Car(x).Move Car(x).Left - 5, Car(x).Top
            End If
        Else
            Car(x).Move Car(x).Left - 5, Car(x).Top
        End If
    End If
lSkip:
    If Car(x).Left <= 24 Then
        Car(x).Left = 672
    End If
Next x

'###############################################################
For x = 0 To nMaxCarRight
    With aiCarR(x)
        .Left = rCar(x).Left
        .Top = rCar(x).Top
        .Right = .Left + 32
        .Bottom = .Top + 32
    End With
Next x
For x = 0 To nMaxCarRight
    If x = 0 Then
        If IntersectRect(IntersectArea, aiCarR(x), aiCarR(nMaxCarRight)) Then
            aiMoveR(x) = False
        Else
            aiMoveR(x) = True
        End If
    Else
        If IntersectRect(IntersectArea, aiCarR(x), aiCarR(x - 1)) Then
            aiMoveR(x) = False
        Else
            aiMoveR(x) = True
        End If
    End If
Next x
For x = 0 To nMaxCarRight
    With aiCarR(x)
        .Left = rCar(x).Left + 22
        .Top = rCar(x).Top
        .Right = .Left + 10
        .Bottom = .Top + 32
    End With
Next x
For x = 0 To nMaxCarRight
    If aiMoveR(x) = True Then
        If IntersectRect(IntersectArea, aiCarR(x), rdLine1) Then
            If lStage > 1 Then 'Green
                'Check for collision with other traffic going
                'by simulating the vehicle has moved left and checking
                'a large area around the vehicle to make sure it has enough
                'room to advance without collision
                With aiCarR(x)
                    .Left = rCar(x).Left + 15
                    .Top = rCar(x).Top - 52
                    .Right = .Left + 32
                    .Bottom = .Top + 110
                End With
                For y = 0 To nMaxCarUp
                    If IntersectRect(IntersectArea, aiCarR(x), aiCarD(y)) Then
                        'if it would collide with a car going in that direction
                        GoTo rSkip
                    End If
                Next y
                rCar(x).Move rCar(x).Left + 5, rCar(x).Top
            End If
        Else
            rCar(x).Move rCar(x).Left + 5, rCar(x).Top
        End If
    End If
rSkip:
    If rCar(x).Left >= 672 Then
        rCar(x).Left = 24
    End If
Next x

'###############################################################

For x = 0 To nMaxCarDown
    With aiCarD(x)
        .Left = dCar(x).Left
        .Top = dCar(x).Top
        .Right = .Left + 32
        .Bottom = .Top + 32
    End With
Next x
For x = 0 To nMaxCarDown
    If x = 0 Then
        If IntersectRect(IntersectArea, aiCarD(x), aiCarD(nMaxCarDown)) Then
            aiMoveD(x) = False
        Else
            aiMoveD(x) = True
        End If
    Else
        If IntersectRect(IntersectArea, aiCarD(x), aiCarD(x - 1)) Then
            aiMoveD(x) = False
        Else
            aiMoveD(x) = True
        End If
    End If
Next x
For x = 0 To nMaxCarDown
    With aiCarD(x)
        .Left = dCar(x).Left
        .Top = dCar(x).Top + 22
        .Right = .Left + 32
        .Bottom = .Top + 10
    End With
Next x
For x = 0 To nMaxCarDown
    If aiMoveD(x) = True Then
        If IntersectRect(IntersectArea, aiCarD(x), rdLine2) Then
            If lStageU > 1 Then 'Green
                'Check for collision with other traffic going left
                'by simulating the vehicle has moved forward and checking
                'a large area around the vehicle to make sure it has enough
                'room to advance without collision
                With aiCarD(x)
                    .Left = dCar(x).Left - 25
                    .Top = dCar(x).Top + 15
                    .Right = .Left + 125
                    .Bottom = .Top + 47
                End With
                For y = 0 To nMaxCarLeft
                    If IntersectRect(IntersectArea, aiCarD(x), aiCar(y)) Then
                        'if it would collide with a car going in that direction
                        GoTo uSkip
                    End If
                    If IntersectRect(IntersectArea, aiCarD(x), aiCarR(y)) Then
                        GoTo uSkip
                    End If
                Next y
                dCar(x).Move dCar(x).Left, dCar(x).Top + 5
            End If
        Else
            dCar(x).Move dCar(x).Left, dCar(x).Top + 5
        End If
    End If
uSkip:
    If dCar(x).Top >= 528 Then
        dCar(x).Top = 16
    End If
Next x
End Sub

Private Sub optLight_Click(Index As Integer)
If optLight(0).Value = True Then
    tmrLight.Enabled = False
    lStage = 1
    LightChange (lStage)
    lStageU = 2
    LightChangeU (lStageU)
    optLightU(1).Value = True
ElseIf optLight(1).Value = True Then
    tmrLight.Enabled = False
    lStage = 2
    LightChange (lStage)
    lStageU = 1
    LightChangeU (lStageU)
    optLightU(0).Value = True
ElseIf optLight(2).Value = True Then
    tmrLight.Enabled = True
    optLightU(2).Value = True
End If
End Sub

Private Sub optLightU_Click(Index As Integer)
If optLightU(0).Value = True Then
    tmrLight.Enabled = False
    lStageU = 1
    LightChangeU (lStageU)
    lStage = 2
    LightChange (lStage)
    optLight(1).Value = True
ElseIf optLightU(1).Value = True Then
    tmrLight.Enabled = False
    lStageU = 2
    LightChangeU (lStageU)
    lStage = 1
    LightChange (lStage)
    optLight(0).Value = True
ElseIf optLightU(2).Value = True Then
    tmrLight.Enabled = True
    optLight(2).Value = True
End If
End Sub

Private Sub tmrLight_Timer()
If lStage = 1 Then
    lStage = 2
    LightChange (lStage)
    lStageU = 1
    LightChangeU (lStageU)
    Exit Sub
ElseIf lStage = 2 Then
    lStage = 1
    LightChange (lStage)
    lStageU = 2
    LightChangeU (lStageU)
    Exit Sub
End If
End Sub

