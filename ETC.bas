Attribute VB_Name = "ETC"
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Const nMaxCarLeft As Single = 6
Public Const nMaxCarRight As Single = 6
Public Const nMaxCarUp As Single = 4
Public Const nMaxCarDown As Single = 5
Public aiCar(10) As RECT
Public aiCarR(10) As RECT
Public aiCarU(10) As RECT
Public aiCarD(10) As RECT
Public rdLine As RECT
Public rdLine1 As RECT
Public rdLine2 As RECT
Public rdLine3 As RECT
Public IntersectArea As RECT
Public aiMove(10) As Boolean
Public aiMoveR(10) As Boolean
Public aiMoveU(10) As Boolean
Public aiMoveD(10) As Boolean
Public nCount As Single
Public lStage As Single
Public lStageU As Single

Public Sub LightChange(nVal As Single)
If nVal = 1 Then
    Main.Red.BackColor = vbRed
    Main.Green.BackColor = vbBlack
ElseIf nVal = 2 Then
    Main.Red.BackColor = vbBlack
    Main.Green.BackColor = vbGreen
End If
End Sub

Public Sub LightChangeU(nVal As Single)
If nVal = 1 Then
    Main.RedU.BackColor = vbRed
    Main.GreenU.BackColor = vbBlack
ElseIf nVal = 2 Then
    Main.RedU.BackColor = vbBlack
    Main.GreenU.BackColor = vbGreen
End If
End Sub
