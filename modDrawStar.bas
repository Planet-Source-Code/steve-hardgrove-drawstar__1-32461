Attribute VB_Name = "modDrawStar"
Option Explicit
'
'This module requires modGeometry
'
Private Declare Function CreatePolygonRgn& Lib "gdi32.dll" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long)
Private Declare Function CreateSolidBrush& Lib "gdi32.dll" (ByVal crColor As Long)
Private Declare Function DeleteObject& Lib "gdi32.dll" (ByVal hObject As Long)
Private Declare Function FrameRgn& Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long)
Private Declare Function PaintRgn& Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long)

Private Const WINDING& = 2

Private Type POINTAPI
  X As Long
  Y As Long
End Type
Public Sub DrawStar(ByRef roObject As Object, ByVal vlX As Long, ByVal vlY As Long, ByVal vlRadius As Long, Optional ByVal olColor As Long = 0, Optional ByVal oeAngle As Double = 0, Optional ByVal oyFrame As Boolean = False, Optional ByVal olFrameWidth As Long = 30, Optional ByVal olFrameColor As Long = 0)

  Dim liFillStyle As Integer
  Dim llBrush As Long
  Dim llFillColor As Long
  Dim llhDC As Long
  Dim llRegion As Long
  Dim llRet As Long
  Dim luPoints() As POINTAPI

  On Error GoTo Error_Handler
  With roObject
    'Remember the object's properties so we can reset them later
    llFillColor = .FillColor
    liFillStyle = .FillStyle
    'Get the handle to the device context
    llhDC = .hdc
    'Set the object's properties
    .FillColor = olColor
    .FillStyle = vbFSSolid
    'Scale the input parameters to whatever scalemode the object is using
    vlX = .ScaleX(vlX, .ScaleMode, vbPixels)
    vlY = .ScaleY(vlY, .ScaleMode, vbPixels)
    vlRadius = .ScaleX(vlRadius, .ScaleMode, vbPixels)
    olFrameWidth = .ScaleX(olFrameWidth, .ScaleMode, vbPixels)
  End With
  'Make sure the angle is between 0° and 360°
  If oeAngle < 0 Or oeAngle > 360 Then oeAngle = 0
  'Calculate the coordinates for the points in the polygon region
  luPoints = GetPoints(vlX, vlY, vlRadius, oeAngle)
  'Create the polygon region representing the star
  llRegion = CreatePolygonRgn(luPoints(0), 11, WINDING)
  'Paint the star
  llRet = PaintRgn(llhDC, llRegion)
  If oyFrame Then
    'Create a brush to draw the frame around the star
    llBrush = CreateSolidBrush(olFrameColor)
    'Paint the frame around the star
    llRet = FrameRgn(llhDC, llRegion, llBrush, olFrameWidth, olFrameWidth)
    'Destroy the brush
    llRet = DeleteObject(llBrush)
  End If
  'Destroy the polygon region
  llRet = DeleteObject(llRegion)
  'Return the object's properties to their original settings
  With roObject
    .FillColor = llFillColor
    .FillStyle = liFillStyle
  End With
  'Erase the array
  Erase luPoints
  Exit Sub

Error_Handler:

End Sub
Private Function GetPoints(ByVal vlX As Long, ByVal vlY As Long, ByVal vlRadius As Long, ByVal veAngle As Double) As POINTAPI()

  Dim llI As Long
  Dim luLineDbl(4) As LineDbl
  Dim luPointDbl(10) As PointDbl
  Dim luPoints(10) As POINTAPI

  On Error GoTo Error_Handler
  'Make the coordinates for the center point of the star the first element of the PointDbl array
  With luPointDbl(0)
    .X = vlX
    .Y = vlY
  End With
  'Add the coordinates for the five points of the star to the PointDbl array
  For llI = 1 To 9 Step 2
    luPointDbl(llI) = GetEndPoints(luPointDbl(0), veAngle, vlRadius)
    veAngle = veAngle + 72
    If veAngle > 360 Then veAngle = veAngle - 360
  Next llI
  'Make some lines based on the five points of the star so we can calculate
  'where they intersect to get the other five points needed to draw the star
  luLineDbl(0).ptStart = luPointDbl(1): luLineDbl(0).ptEnd = luPointDbl(5)
  luLineDbl(1).ptStart = luPointDbl(3): luLineDbl(1).ptEnd = luPointDbl(7)
  luLineDbl(2).ptStart = luPointDbl(5): luLineDbl(2).ptEnd = luPointDbl(9)
  luLineDbl(3).ptStart = luPointDbl(7): luLineDbl(3).ptEnd = luPointDbl(1)
  luLineDbl(4).ptStart = luPointDbl(9): luLineDbl(4).ptEnd = luPointDbl(3)
  'Add the intersection points of the lines to the PointDbl array
  Call LineIntersect(luLineDbl(0), luLineDbl(4), luPointDbl(2))
  Call LineIntersect(luLineDbl(0), luLineDbl(1), luPointDbl(4))
  Call LineIntersect(luLineDbl(1), luLineDbl(2), luPointDbl(6))
  Call LineIntersect(luLineDbl(2), luLineDbl(3), luPointDbl(8))
  Call LineIntersect(luLineDbl(3), luLineDbl(4), luPointDbl(10))
  'Add the ten points of the star to the POINTAPI array
  For llI = 1 To 10
    With luPoints(llI)
      .X = CLng(luPointDbl(llI).X)
      .Y = CLng(luPointDbl(llI).Y)
    End With
  Next llI
  'Make the first point in the POINTAPI array the same as the last so the polygon region will be closed
  With luPoints(0)
    .X = luPoints(10).X
    .Y = luPoints(10).Y
  End With
  'Set the return value
  GetPoints = luPoints
  'Erase the arrays
  Erase luLineDbl
  Erase luPointDbl
  Erase luPoints
  Exit Function

Error_Handler:

End Function
