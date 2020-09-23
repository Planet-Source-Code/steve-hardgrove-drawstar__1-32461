VERSION 5.00
Begin VB.Form frmStarAngle 
   AutoRedraw      =   -1  'True
   Caption         =   "Star Angle"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStarAngle.frx":0000
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmStarAngle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private meAngle As Double
Private meSize As Double
Private muCenterPoint As PointDbl

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim luPoint As PointDbl

  On Error Resume Next
  With luPoint
    .X = X
    .Y = Y
  End With
  If Distance(muCenterPoint, luPoint) <= meSize * 1.1 Then
    Me.MousePointer = vbCrosshair
  Else
    Me.MousePointer = vbDefault
  End If
  DoEvents

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim luLine As LineDbl
  Dim luPoint As PointDbl

  On Error Resume Next
  With luPoint
    .X = X
    .Y = Y
  End With
  If Distance(muCenterPoint, luPoint) <= meSize * 1.1 Then
    With luLine
      .ptStart = muCenterPoint
      .ptEnd = luPoint
    End With
    meAngle = CDbl(CLng(LineAngleDegrees(luLine)))
    If meAngle = 360 Then meAngle = 0
    If Button = vbRightButton Then
      Select Case meAngle
        Case 0 To 179
          meAngle = meAngle + 180
        Case Else
          meAngle = meAngle - 180
      End Select
    End If
    Call Form_Resize
  End If

End Sub
Private Sub Form_Resize()

  Dim leI As Double
  Dim luPoint1 As PointDbl
  Dim luPoint2 As PointDbl

  On Error Resume Next
  With Me
    If .WindowState <> vbMinimized Then
      .Cls
      With muCenterPoint
        .X = Me.ScaleWidth / 2
        .Y = Me.ScaleHeight / 2
        If .X > .Y Then
          meSize = .Y * 0.9
        Else
          meSize = .X * 0.9
        End If
      End With
      Call DrawStar(Me, muCenterPoint.X, muCenterPoint.Y, meSize, &H101E0, meAngle, True, .ScaleX(30, vbTwips, .ScaleMode), &H1EFFF)
      Me.Circle (muCenterPoint.X, muCenterPoint.Y), meSize * 1.1
      For leI = 0 To 359
        If leI Mod 5 = 0 Then
          If leI Mod 15 = 0 Then
            If leI Mod 30 = 0 Then
              luPoint1 = GetEndPoints(muCenterPoint, leI, meSize)
              Select Case leI
                Case 0
                  .CurrentX = luPoint1.X - (.TextWidth(CStr(leI)) / 2)
                  .CurrentY = luPoint1.Y - (.TextHeight(CStr(leI)) / 2) + .ScaleX(60, vbTwips, .ScaleMode)
                Case Is < 180
                  .CurrentX = luPoint1.X - .TextWidth(CStr(leI))
                  .CurrentY = luPoint1.Y - (.TextHeight(CStr(leI)) / 2)
                Case 180
                  .CurrentX = luPoint1.X - (.TextWidth(CStr(leI)) / 2)
                  .CurrentY = luPoint1.Y - (.TextHeight(CStr(leI)) / 2) - .ScaleX(60, vbTwips, .ScaleMode)
                Case Else
                  .CurrentX = luPoint1.X
                  .CurrentY = luPoint1.Y - (.TextHeight(CStr(leI)) / 2)
              End Select
              Me.Print CStr(leI)
            End If
            luPoint1 = GetEndPoints(muCenterPoint, leI, meSize * 1.03)
          Else
            luPoint1 = GetEndPoints(muCenterPoint, leI, meSize * 1.05)
          End If
        Else
          luPoint1 = GetEndPoints(muCenterPoint, leI, meSize * 1.07)
        End If
        luPoint2 = GetEndPoints(muCenterPoint, leI, meSize * 1.1)
        If leI = meAngle Then
          Me.Line (luPoint1.X, luPoint1.Y)-(luPoint2.X, luPoint2.Y), vbRed
        Else
          Me.Line (luPoint1.X, luPoint1.Y)-(luPoint2.X, luPoint2.Y), vbBlack
        End If
      Next leI
      .CurrentX = .ScaleX(60, vbTwips, .ScaleMode)
      .CurrentY = .ScaleY(60, vbTwips, .ScaleMode)
      Me.Print CStr(meAngle) & Chr$(176)
      .Refresh
    End If
  End With

End Sub
