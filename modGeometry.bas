Attribute VB_Name = "modGeometry"
Option Explicit

'Some of the functions in this module were borrowed from
'various other modules downloaded from VB websites

Public Type PointDbl   'Point structure (in Doubles)
  X As Double          'X-coordinate of point.
  Y As Double          'Y-coordinate of point.
End Type

Public Type LineDbl    'Line structure (in Doubles)
  ptStart As PointDbl  'Starting point (X, Y) on line.
  ptEnd As PointDbl    'Ending point (X, Y) on line.
End Type

Public Type ArcStruct
  bValidArc As Boolean 'Is this a valid arc.
  ptStart As PointDbl  'Starting point.
  PtMid As PointDbl    'Mid point.
  ptEnd As PointDbl    'Ending point.
  ptCenter As PointDbl 'Center point.
  dRadius As Double    'Radius.
  dRadsStart As Double 'Starting angle in radians.
  dRadsMid As Double   'Mid angle in radians.
  dRadsEnd As Double   'Ending angle in radians.
End Type

Private Const mcPI# = 3.14159265358979    'Pi
Public Function AreaOfCircle(ByVal veRadius As Double) As Double

  On Error GoTo Error_Handler
  AreaOfCircle = mcPI * (veRadius ^ 2)
  Exit Function

Error_Handler:

End Function
Public Function AreaOfRect(ByVal veLength As Double, ByVal veWidth As Double) As Double

  On Error GoTo Error_Handler
  AreaOfRect = veLength * veWidth
  Exit Function

Error_Handler:

End Function
Public Function AreaOfRing(ByVal veInnerRadius As Double, ByVal veOuterRadius As Double) As Double

  On Error GoTo Error_Handler
  AreaOfRing = AreaOfCircle(veOuterRadius) - AreaOfCircle(veInnerRadius)
  Exit Function

Error_Handler:

End Function
Public Function AreaOfSphere(ByVal veRadius As Double) As Double

  On Error GoTo Error_Handler
  AreaOfSphere = 4 * mcPI * (veRadius ^ 2)
  Exit Function

Error_Handler:

End Function
Public Function AreaOfSquare(ByVal veSide As Double) As Double

  On Error GoTo Error_Handler
  AreaOfSquare = veSide * veSide
  Exit Function

Error_Handler:

End Function
Public Function AreaOfSquare2(ByVal veDiagonal As Double) As Double

  On Error GoTo Error_Handler
  AreaOfSquare2 = veDiagonal * veDiagonal / 2
  Exit Function

Error_Handler:

End Function
Public Function AreaOfTrap(ByVal vePerpendicularHeight As Double, ByVal veParallelSide1 As Double, ByVal veParallelSide2 As Double) As Double

  On Error GoTo Error_Handler
  AreaOfTrap = vePerpendicularHeight * (veParallelSide1 + veParallelSide2) / 2
  Exit Function

Error_Handler:

End Function
Public Function AreaOfTriangle(ByVal veSide As Double, ByVal vePerpendicularHeight As Double) As Double

  On Error GoTo Error_Handler
  AreaOfTriangle = veSide * vePerpendicularHeight / 2
  Exit Function

Error_Handler:

End Function
Public Function AreaOfTriangle2(ByVal veSide1 As Double, ByVal veSide2 As Double, ByVal veSide3 As Double) As Double

  Dim leCosC As Double

  On Error GoTo Error_Handler
  leCosC = (veSide1 * veSide1 + veSide2 * veSide2 - veSide3 * veSide3) / (2 * veSide1 * veSide2)
  AreaOfTriangle2 = veSide1 * veSide2 * Sqr(1 - leCosC * leCosC) / 2
  Exit Function

Error_Handler:

End Function
Public Function CalcArc(ByRef ruPoint1 As PointDbl, ByRef ruPoint2 As PointDbl, ByRef ruPoint3 As PointDbl) As ArcStruct

  'Calculates all data needed to draw an arc from 3 points.
  'Returns an ArcStruct structure. (see declares section)

  'Example Syntax:
  'Dim Arc1 As ArcStruct
  'Arc1 = CalcArc(ruPoint1, ruPoint2, ruPoint3)
  'If Arc1.bValidArc Then
  '  With Arc1
  '    picBox.Circle(.luCenter.X, .luCenter.Y), .dRadius, lAnyColor, .dRadsStart, .dRadsEnd
  '  End With
  'End If

  Dim leRads(3) As Double
  Dim luLine(3) As LineDbl
  Dim luCenter As PointDbl

  On Error GoTo Error_Handler
  'Setup 2 lines using the 3 points.
  luLine(0).ptStart = ruPoint1:  luLine(0).ptEnd = ruPoint2
  luLine(1).ptStart = ruPoint2:  luLine(1).ptEnd = ruPoint3
  'Create a perpendicular line from the
  'centers of each of the two lines.
  luLine(2) = PerpLineCenter(luLine(0))
  luLine(3) = PerpLineCenter(luLine(1))
  'If the perp lines don't intersect then the 3 points
  'are on a straight line and cannot be an arc.
  If LineIntersect(luLine(2), luLine(3), luCenter) <> -1 Then
    'If the perp lines intersect then it forms an arc.
    'Setup 3 lines from the center; 1 line to each outer point.
    luLine(0).ptStart = luCenter:    luLine(0).ptEnd = ruPoint1
    luLine(1).ptStart = luCenter:    luLine(1).ptEnd = ruPoint2
    luLine(2).ptStart = luCenter:    luLine(2).ptEnd = ruPoint3
    leRads(0) = LineAngleRadians(luLine(0))
    leRads(1) = LineAngleRadians(luLine(1))
    leRads(2) = LineAngleRadians(luLine(2))
    'An arc is always drawn counter-clockwise, so order the points.
    If Not IsBetween(leRads(1), leRads(0), leRads(2), False) Then
      'leRads(1) is not between leRads(0) and leRads(2),
      'so the arc must wrap around the 0° mark. This means the
      'greater of leRads(0) and leRads(2) is the start point.
      If leRads(2) > leRads(0) Then 'Reversed, so swap points.
        leRads(3) = leRads(0)
        luLine(3) = luLine(0)
        leRads(0) = leRads(2)
        luLine(0) = luLine(2)
        leRads(2) = leRads(3)
        luLine(2) = luLine(3)
      End If
    Else
      'No wrap around, so the lessor of leRads(0)
      'and leRads(2) is the start point.
      If leRads(2) < leRads(0) Then 'Reversed, so swap points.
        leRads(3) = leRads(0)
        luLine(3) = luLine(0)
        leRads(0) = leRads(2)
        luLine(0) = luLine(2)
        leRads(2) = leRads(3)
        luLine(2) = luLine(3)
      End If
    End If
    'Now that the points and angles are all in order, return the data.
    With CalcArc
      .bValidArc = True
      .ptStart = luLine(0).ptEnd
      .PtMid = luLine(1).ptEnd
      .ptEnd = luLine(2).ptEnd
      .ptCenter = luCenter
      .dRadius = Distance(.ptCenter, .ptStart)
      .dRadsStart = leRads(0)
      .dRadsMid = leRads(1)
      .dRadsEnd = leRads(2)
    End With
  Else
    'Straight line; Set bValidArc to False.
    CalcArc.bValidArc = False
  End If
  Exit Function

Error_Handler:
  CalcArc.bValidArc = False

End Function
Public Function CircumferenceOfCircle(ByVal veRadius As Double) As Double

  On Error GoTo Error_Handler
  CircumferenceOfCircle = 2 * mcPI * veRadius
  Exit Function

Error_Handler:

End Function
Public Function DegreesToRadians(ByVal veDegrees As Double) As Double

  'Converts Degrees to Radians.

  On Error GoTo Error_Handler
  DegreesToRadians = veDegrees * (mcPI / 180#)
  Exit Function

Error_Handler:

End Function
Public Function DiagonalOfRectangle(ByVal veLength As Double, ByVal veWidth As Double) As Double

  On Error GoTo Error_Handler
  DiagonalOfRectangle = Sqr(veWidth * veWidth + veLength * veLength)
  Exit Function

Error_Handler:

End Function
Public Function DiagonalOfSquare(ByVal veSide As Double) As Double

  On Error GoTo Error_Handler
  DiagonalOfSquare = veSide * Sqr(2)
  Exit Function

Error_Handler:

End Function
Public Function Distance(ByRef ruStartPoint As PointDbl, ByRef ruEndPoint As PointDbl) As Double

  'Calculates the distance between 2 points.

  On Error GoTo Error_Handler
  'Standard hypotenuse equation (c = Sqr(a^2 + b^2))
  Distance = Sqr(((ruEndPoint.X - ruStartPoint.X) ^ 2) + ((ruEndPoint.Y - ruStartPoint.Y) ^ 2))
  Exit Function

Error_Handler:

End Function
Public Function Div(ByVal veNumer As Double, ByVal veDenom As Double) As Double

  ' Divides 2 numbers avoiding a "Division by zero" error.

  On Error GoTo Error_Handler
  If veDenom <> 0 Then
    Div = veNumer / veDenom
  Else
    Div = 0
  End If
  Exit Function

Error_Handler:

End Function
Public Function GetEndPoints(ByRef ruStartPoint As PointDbl, ByVal veAngle As Double, ByVal veLength As Double, Optional ByVal oy0DegNorth As Boolean = True) As PointDbl

  'Returns the end coordinates of a line where
  'the starting coordinates, angle, and length are known

  Dim leRadians As Double

  On Error GoTo Error_Handler
  'Set the value for radians
  leRadians = mcPI / 180
  'Make sure the angle is between 0 and 360
  If veAngle < 0 Then veAngle = 0
  If veAngle > 360 Then veAngle = 360
  'Correct the angle so 0 is north, 180 is south, etc.
  'There may be a better way to do this, but it works
  If oy0DegNorth Then
    Select Case veAngle
      Case 0 To 180, 360
        veAngle = Abs(180 - veAngle)
      Case 180 To 360
        veAngle = Abs(540 - veAngle)
    End Select
  End If
  'Set the X and Y values
  With GetEndPoints
    .X = ruStartPoint.X + Sin(veAngle * leRadians) * veLength
    .Y = ruStartPoint.Y + Cos(veAngle * leRadians) * veLength
  End With
  Exit Function

Error_Handler:

End Function
Public Function IsBetween(ByVal vvTestData As Variant, ByVal vvLowerBound As Variant, ByVal vvUpperBound As Variant, Optional ByVal oyInclusive As Boolean = True) As Boolean

  'Returns True if vvTestData is between vvLowerBound and vvUpperBound.
  'oyInclusive = Are the bounds included in the test?

  Dim lvTemp As Variant

  On Error GoTo Error_Handler
  If vvLowerBound = vvUpperBound Then
    Exit Function   'Returns false if upper and lower bounds are equal.
  Else
    If vvLowerBound > vvUpperBound Then
      'If bounds are reversed, swap them.
      lvTemp = vvLowerBound
      vvLowerBound = vvUpperBound
      vvUpperBound = lvTemp
    End If
    If oyInclusive Then
      'If bounds are included in test (use >= and <=).
      IsBetween = (vvTestData >= vvLowerBound) And (vvTestData <= vvUpperBound)
    Else
      'If bounds are not included in test (use > and <).
      IsBetween = (vvTestData > vvLowerBound) And (vvTestData < vvUpperBound)
    End If
  End If
  Exit Function

Error_Handler:

End Function
Public Function LineAngleDegrees(ByRef ruLine As LineDbl, Optional ByVal oy0DegNorth As Boolean = True) As Double

  'Returns the angle of a line in degrees (see LineAngleRadians).

  Dim leAngle As Double

  On Error GoTo Error_Handler
  leAngle = RadiansToDegrees(LineAngleRadians(ruLine))
  'Make sure the angle is between 0 and 360
  If leAngle < 0 Then leAngle = 0
  If leAngle > 360 Then leAngle = 360
  'Correct the angle so 0 is north, 180 is south, etc.
  'There may be a better way to do this, but it works
  If oy0DegNorth Then
    Select Case leAngle
      Case 360
        leAngle = leAngle - 270
      Case 0 To 90
        leAngle = Abs(90 - leAngle)
      Case 90 To 360
        leAngle = Abs(450 - leAngle)
    End Select
  End If
  LineAngleDegrees = Format(leAngle, "##0.0##")
  Exit Function

Error_Handler:

End Function
Public Function LineAngleRadians(ByRef ruLine As LineDbl) As Double

  'Calculates the angle(in radians) of a line from ptStart to ptEnd.

  Dim leAngle As Double
  Dim leDeltaX As Double
  Dim leDeltaY As Double

  On Error GoTo Error_Handler
  With ruLine
    leDeltaX = .ptEnd.X - .ptStart.X
    leDeltaY = .ptEnd.Y - .ptStart.Y
  End With
  If leDeltaX = 0 Then      'Vertical
    If leDeltaY < 0 Then
      leAngle = mcPI / 2
    Else
      leAngle = mcPI * 1.5
    End If
  ElseIf leDeltaY = 0 Then  'Horizontal
    If leDeltaX >= 0 Then
      leAngle = 0
    Else
      leAngle = mcPI
    End If
  Else    'Angled
    'Note: ++ = positive X, positive Y; +- = positive X, negative Y; etc.
    'On a true coordinate plane, Y increases as it move upward.
    'In VB coordinates, Y is reversed. It increases as it moves downward.
    'Calc for true Upper Right Quadrant (++) (For VB this is +-)
    leAngle = Atn(Abs(leDeltaY / leDeltaX))        'VB Upper Right (+-)
    'Correct for other 3 quadrants in VB coordinates (Reversed Y)
    If leDeltaX >= 0 And leDeltaY >= 0 Then       'VB Lower Right (++)
      leAngle = (mcPI * 2) - leAngle
    ElseIf leDeltaX < 0 And leDeltaY >= 0 Then    'VB Lower Left (-+)
      leAngle = mcPI + leAngle
    ElseIf leDeltaX < 0 And leDeltaY < 0 Then     'VB Upper Left (--)
      leAngle = mcPI - leAngle
    End If
  End If
  LineAngleRadians = leAngle
  Exit Function

Error_Handler:

End Function
Public Function LineIntersect(ByRef ruLine1 As LineDbl, ByRef ruLine2 As LineDbl, ByRef ruIntersect As PointDbl) As Integer

  'Calculate the intersection point of any two given non-parallel lines.
  '
  'Returns:  -1 = lines are parallel (no intersection).
  '           0 = Neither line contains the intersect point between its points.**
  '           1 = ruLine1 contains the intersect point between its points.**
  '           2 = ruLine2 contains the intersect point between its points.**
  '           3 = Both Lines contain the intersect point between their points.**
  '           ** Lines Do intersect; Also fills in the ruIntersect point.
  '
  'BTW:       There are 18 lines of pure code, 25 lines of pure comments and 6
  '           mixed lines in this function, just in case you were wondering. (:oþ}

  Dim leDenom As Double
  Dim lePctDelta1 As Double
  Dim lePctDelta2 As Double
  Dim liReturn As Integer
  Dim luDelta(2) As PointDbl
  Dim lyIntersect As Boolean

  On Error GoTo Error_Handler
  'Calculate the Deltas (distance of X2 - X1 or Y2 - Y1 of any 2 points)
  luDelta(0).X = ruLine1.ptStart.X - ruLine2.ptStart.X   'ruLine1-ruLine2.ptStart X-Cross-luDelta
  luDelta(0).Y = ruLine1.ptStart.Y - ruLine2.ptStart.Y   'ruLine1-ruLine2.ptStart Y-Cross-luDelta
  luDelta(1).X = ruLine1.ptEnd.X - ruLine1.ptStart.X   'ruLine1 X-luDelta
  luDelta(1).Y = ruLine1.ptEnd.Y - ruLine1.ptStart.Y   'ruLine1 Y-luDelta
  luDelta(2).X = ruLine2.ptEnd.X - ruLine2.ptStart.X   'ruLine2 X-luDelta
  luDelta(2).Y = ruLine2.ptEnd.Y - ruLine2.ptStart.Y   'ruLine2 Y-luDelta
  'Calculate the denominator (zero = parallel (no intersection))
  'Formula: (L2Dy * L1Dx) - (L2Dx * L1Dy)
  liReturn = -1
  leDenom = (luDelta(2).Y * luDelta(1).X) - (luDelta(2).X * luDelta(1).Y)
  lyIntersect = (leDenom <> 0)
  If lyIntersect Then
    'The lines will intersect somewhere.
    'Solve for both lines using the Cross-Deltas (luDelta(0))
    'This yields percentage (0.1 = 10%; 1 = 100%) of the distance
    'between ptStart and ptEnd, of the opposite line, where the line used
    'in the calculation will cross it.
    '0 = ptStart direct hit; 1 = ptEnd direct hit; 0.5 = Centered between Pts; etc.
    'If < 0 or > 1 then the lines still intersect, just not between the points.
    'Solve for ruLine1 where ruLine2 will cross it.
    lePctDelta1 = ((luDelta(2).X * luDelta(0).Y) - (luDelta(2).Y * luDelta(0).X)) / leDenom
    'Solve for ruLine2 where ruLine1 will cross it.
    lePctDelta2 = ((luDelta(1).X * luDelta(0).Y) - (luDelta(1).Y * luDelta(0).X)) / leDenom
    'Check for absolute intersection. If the percentage is not between
    '0 and 1 then the lines will not intersect between their points.
    'Returns 0, 1, 2 or 3.
    liReturn = IIf(IsBetween(lePctDelta1, 0#, 1#), 1, 0) Or IIf(IsBetween(lePctDelta2, 0#, 1#), 2, 0)
    'Calculate point of intersection on ruLine1 and fill ruIntersect.
    ruIntersect.X = ruLine1.ptStart.X + (lePctDelta1 * luDelta(1).X)
    ruIntersect.Y = ruLine1.ptStart.Y + (lePctDelta1 * luDelta(1).Y)
  End If
  'Return the results.
  LineIntersect = liReturn
  Exit Function

Error_Handler:

End Function
Public Function PerpLineCenter(ByRef ruLine As LineDbl) As LineDbl

  'Returns a line perpendicular (90°) to ruLine1 using
  'the center of ruLine1 as the first point.

  Dim leDeltaX As Double
  Dim leDeltaY As Double
  Dim luLine As LineDbl

  On Error GoTo Error_Handler
  With luLine
    .ptStart.X = (ruLine.ptStart.X + ruLine.ptEnd.X) / 2#
    .ptStart.Y = (ruLine.ptStart.Y + ruLine.ptEnd.Y) / 2#
    leDeltaX = .ptStart.X - ruLine.ptStart.X
    leDeltaY = .ptStart.Y - ruLine.ptStart.Y
    .ptEnd.X = .ptStart.X + -leDeltaY
    .ptEnd.Y = .ptStart.Y + leDeltaX
  End With
  PerpLineCenter = luLine
  Exit Function

Error_Handler:

End Function
Public Function PointOnLine(ByRef ruStartPoint As PointDbl, ByRef ruEndPoint As PointDbl, ByVal veDistance As Double) As PointDbl

  'Returns a point on a line at veDistance from ruStart.
  'This point need not be between ruStart and ruEnd.

  Dim lgDX As Single
  Dim lgDY As Single
  Dim lgLen As Single
  Dim lgPct As Single

  On Error GoTo Error_Handler
  If veDistance > 1000000 Then
    veDistance = 1000000
  End If
  lgLen = Distance(ruStartPoint, ruEndPoint)
  With ruStartPoint
    If lgLen > 0 Then
      lgDX = ruEndPoint.X - .X
      lgDY = ruEndPoint.Y - .Y
      lgPct = Div(veDistance, lgLen)
      PointOnLine.X = .X + (lgDX * lgPct)
      PointOnLine.Y = .Y + (lgDY * lgPct)
    Else
      PointOnLine.X = .X
      PointOnLine.Y = .Y
    End If
  End With
  Exit Function

Error_Handler:

End Function
Public Function RadiansToDegrees(ByVal veRadians As Double) As Double

  'Converts Radians to Degrees.

  On Error GoTo Error_Handler
  RadiansToDegrees = veRadians * (180# / mcPI)
  Exit Function

Error_Handler:

End Function
Public Function RadiusOfCircle(ByVal veCircumference As Double) As Double

  On Error GoTo Error_Handler
  RadiusOfCircle = veCircumference / mcPI / 2#
  Exit Function

Error_Handler:

End Function
Function VolumeOfCone(ByVal veHeight As Double, ByVal veRadius As Double) As Double

  On Error GoTo Error_Handler
  VolumeOfCone = veHeight * (veRadius ^ 2) * mcPI / 3#
  Exit Function

Error_Handler:

End Function
Public Function VolumeOfCylinder(ByVal veHeight As Double, ByVal veRadius As Double) As Double

  On Error GoTo Error_Handler
  VolumeOfCylinder = mcPI * (veRadius ^ 2) * veHeight
  Exit Function

Error_Handler:

End Function
Public Function VolumeOfPipe(ByVal veHeight As Double, ByVal veOuterRadius As Double, ByVal veInnerRadius As Double) As Double

  On Error GoTo Error_Handler
  VolumeOfPipe = VolumeOfCylinder(veHeight, veOuterRadius) - VolumeOfCylinder(veHeight, veInnerRadius)
  Exit Function

Error_Handler:

End Function
Public Function VolumeOfPyramid(ByVal veHeight As Double, ByVal veBaseArea As Double) As Double

  On Error GoTo Error_Handler
  VolumeOfPyramid = veHeight * veBaseArea / 3#
  Exit Function

Error_Handler:

End Function
Public Function VolumeOfSphere(ByVal veRadius As Double) As Double

  On Error GoTo Error_Handler
  VolumeOfSphere = mcPI * (veRadius ^ 3) * 4# / 3#
  Exit Function

Error_Handler:

End Function
Public Function VolumeOfTruncPyramid(ByVal veHeight As Double, ByVal veBaseArea1 As Double, ByVal veBaseArea2 As Double) As Double

  On Error GoTo Error_Handler
  VolumeOfTruncPyramid = veHeight * (veBaseArea1 + veBaseArea2 + Sqr(veBaseArea1) * Sqr(veBaseArea2)) / 3#
  Exit Function

Error_Handler:

End Function
