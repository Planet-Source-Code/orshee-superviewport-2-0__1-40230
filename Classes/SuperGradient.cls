VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SuperGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Property Storage Variables
Private mvarColor1    As Long
Private mvarColor2    As Long
Private mvarAngle     As Single

'Property Default Constants
Private Const mvar_def_Color1   As Long = &HFFFFD0
Private Const mvar_def_Color2   As Long = &H400000
Private Const mvar_def_Angle    As Single = 315

'Misc Constants
Private Const PI    As Double = 3.14159265358979
Private Const RADS  As Double = PI / 180    '<Degrees> * RADS = radians

'TypeDefs
Private Type PointSng   'Internal Point structure
    x   As Single       'Uses Singles for more precision.
    y   As Single
End Type

Private Type PointAPI   'API Point structure
    x   As Long
    y   As Long
End Type


Public Function Draw(hWnd As Long, hDC As Long) As Boolean

'Note: This class uses API functions to draw. If the
'      destination object is in AutoRedraw mode, the
'      Refresh method for that object must be invoked.

    Dim myRet    As Long
    Dim myIdx    As Long
    Dim myRect   As RECT

    On Error GoTo LocalError
    
    'Stop the window from updating until we're finished.
    myRet = LockWindowUpdate(hWnd)
    
    'Get the client rect in pixels
    myRet = GetClientRect(hWnd, myRect)
    
    'Test for possible errors (GetClientRect failure or Rect < 2 pixels)
    If myRet <> 0 Then
        If myRect.Right > 1 And myRect.Bottom > 1 Then
            myIdx = DrawGradient(hDC, myRect.Right, myRect.Bottom)
            Draw = (myIdx > 0)
        End If
    End If
    
NormalExit:
    'Unlock the window to allow it to update now.
    myRet = LockWindowUpdate(0)
    Exit Function
    
LocalError:
    MsgBox Err.Description, vbExclamation
    Resume NormalExit

End Function

Public Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long, ByVal lSteps As Long, laRetColors() As Long) As Long

'Creates an array of colors blending from
'Color1 to Color2 in lSteps number of steps.
'Returns the count and fills the laRetColors() array.

Dim lIdx    As Long
Dim lRed    As Long
Dim lGrn    As Long
Dim lBlu    As Long
Dim fRedStp As Single
Dim fGrnStp As Single
Dim fBluStp As Single

    'Stop possible error
    If lSteps < 2 Then lSteps = 2
    
    'Extract Red, Blue and Green values from the start and end colors.
    lRed = (lColor1 And &HFF&)
    lGrn = (lColor1 And &HFF00&) / &H100
    lBlu = (lColor1 And &HFF0000) / &H10000
    
    'Find the amount of change for each color element per color change.
    fRedStp = Div(CSng((lColor2 And &HFF&) - lRed), CSng(lSteps))
    fGrnStp = Div(CSng(((lColor2 And &HFF00&) / &H100&) - lGrn), CSng(lSteps))
    fBluStp = Div(CSng(((lColor2 And &HFF0000) / &H10000) - lBlu), CSng(lSteps))
    
    'Create the colors
    ReDim laRetColors(lSteps - 1)
    laRetColors(0) = lColor1            'First Color
    laRetColors(lSteps - 1) = lColor2   'Last Color
    For lIdx = 1 To lSteps - 2          'All Colors between
        laRetColors(lIdx) = CLng(lRed + (fRedStp * CSng(lIdx))) + _
            (CLng(lGrn + (fGrnStp * CSng(lIdx))) * &H100&) + _
            (CLng(lBlu + (fBluStp * CSng(lIdx))) * &H10000)
    Next lIdx
    
    'Return number of colors in array
    BlendColors = lSteps

End Function

Private Function DrawGradient(ByVal hDC As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long
    Dim bDone       As Boolean
    Dim iIncX       As Integer
    Dim iIncY       As Integer
    Dim lIdx        As Long
    Dim lRet        As Long
    Dim hPen        As Long
    Dim hOldPen     As Long
    Dim lPointCnt   As Long
    Dim laColors()  As Long
    Dim fMovX       As Single
    Dim fMovY       As Single
    Dim fDist       As Single
    Dim fAngle      As Single
    Dim fLongSide   As Single
    Dim uTmpPt      As POINT
    Dim uaPts()     As POINT
    Dim uaTmpPts()  As PointSng
    
    On Error GoTo LocalError
    
    'Start with center of rect
    ReDim uaTmpPts(2)
    uaTmpPts(2).x = Int(lWidth / 2)
    uaTmpPts(2).y = Int(lHeight / 2)
    
    'Calc distance to furthest edge as if rect were square
    fLongSide = IIf(lWidth > lHeight, lWidth, lHeight)
    fDist = (Sqr((fLongSide ^ 2) + (fLongSide ^ 2)) + 2) / 2
    
    'Create points to the left and the right at a 0º angle (horizontal)
    uaTmpPts(0).x = uaTmpPts(2).x - fDist
    uaTmpPts(0).y = uaTmpPts(2).y
    uaTmpPts(1).x = uaTmpPts(2).x + fDist
    uaTmpPts(1).y = uaTmpPts(2).y
    
    'Lines will be drawn perpendicular to mvarAngle so
    'add 90º and correct for 360º wrap
    fAngle = CDbl(mvarAngle + 90) - Int(Int(CDbl(mvarAngle + 90) / 360#) * 360#)
    
    'Rotate second and third points to fAngle
    Call RotatePoint(uaTmpPts(2), uaTmpPts(0), fAngle)
    Call RotatePoint(uaTmpPts(2), uaTmpPts(1), fAngle)
    
    'We now have a line that crosses the center and
    'two sides of the rect at the correct angle.
    
    'Calc the starting quadrant, direction of and amount of first move
    '(fMovX, fMovY moves line from center to starting edge)
    'and direction of each incremental move (iIncX, iIncY).
    Select Case mvarAngle
        Case 0 To 90
            'Left Bottom
            If Abs(uaTmpPts(0).x - uaTmpPts(1).x) <= Abs(uaTmpPts(0).y - uaTmpPts(1).y) Then
                'Move line to left edge; Draw left to right
                fMovX = IIf(uaTmpPts(0).x > uaTmpPts(1).x, -uaTmpPts(0).x, -uaTmpPts(1).x)
                fMovY = 0
                iIncX = 1
                iIncY = 0
            Else
                'Move line to bottom edge; Draw bottom to top
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).y > uaTmpPts(1).y, lHeight - uaTmpPts(1).y, lHeight - uaTmpPts(0).y)
                iIncX = 0
                iIncY = -1
            End If
        Case 90 To 180
            'Right Bottom
            If Abs(uaTmpPts(0).x - uaTmpPts(1).x) <= Abs(uaTmpPts(0).y - uaTmpPts(1).y) Then
                'Move line to right edge; Draw right to left
                fMovX = IIf(uaTmpPts(0).x > uaTmpPts(1).x, lWidth - uaTmpPts(1).x, lWidth - uaTmpPts(0).x)
                fMovY = 0
                iIncX = -1
                iIncY = 0
            Else
                'Move line to bottom edge; Draw bottom to top
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).y > uaTmpPts(1).y, lHeight - uaTmpPts(1).y, lHeight - uaTmpPts(0).y)
                iIncX = 0
                iIncY = -1
            End If
        Case 180 To 270
            'Right Top
            If Abs(uaTmpPts(0).x - uaTmpPts(1).x) <= Abs(uaTmpPts(0).y - uaTmpPts(1).y) Then
                'Move line to right edge; Draw right to left
                fMovX = IIf(uaTmpPts(0).x > uaTmpPts(1).x, lWidth - uaTmpPts(1).x, lWidth - uaTmpPts(0).x)
                fMovY = 0
                iIncX = -1
                iIncY = 0
            Else
                'Move line to top edge; Draw top to bottom
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).y > uaTmpPts(1).y, -uaTmpPts(0).y, -uaTmpPts(1).y)
                iIncX = 0
                iIncY = 1
            End If
        Case Else   '(270 to 360)
            'Left Top
            If Abs(uaTmpPts(0).x - uaTmpPts(1).x) <= Abs(uaTmpPts(0).y - uaTmpPts(1).y) Then
                'Move line to left edge; Draw left to right
                fMovX = IIf(uaTmpPts(0).x > uaTmpPts(1).x, -uaTmpPts(0).x, -uaTmpPts(1).x)
                fMovY = 0
                iIncX = 1
                iIncY = 0
            Else
                'Move line to top edge; Draw top to bottom
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).y > uaTmpPts(1).y, -uaTmpPts(0).y, -uaTmpPts(1).y)
                iIncX = 0
                iIncY = 1
            End If
    End Select
    
    'At this point we could calculate where the lines will cross the rect edges, but
    'this would slow things down. The picObj clipping region will take care of this.
    
    'Start with 1000 points and add more if needed. This increases
    'speed by not re-dimming the array in each loop.
    ReDim uaPts(999)
    
    'Set the first two points in the array
    uaPts(0).x = uaTmpPts(0).x + fMovX
    uaPts(0).y = uaTmpPts(0).y + fMovY
    uaPts(1).x = uaTmpPts(1).x + fMovX
    uaPts(1).y = uaTmpPts(1).y + fMovY
    
    lIdx = 2
    'Create the rest of the points by incrementing both points
    'on each line iIncX, iIncY from the previous line's points.
    'Where we stop depends on the direction of travel.
    'We'll continue until both points in a set reach the end.
    While Not bDone
        uaPts(lIdx).x = uaPts(lIdx - 2).x + iIncX
        uaPts(lIdx).y = uaPts(lIdx - 2).y + iIncY
        lIdx = lIdx + 1
        Select Case True
            Case iIncX > 0  'Moving Left to Right
                bDone = uaPts(lIdx - 1).x > lWidth And uaPts(lIdx - 2).x > lWidth
            Case iIncX < 0  'Moving Right to Left
                bDone = uaPts(lIdx - 1).x < 0 And uaPts(lIdx - 2).x < 0
            Case iIncY > 0  'Moving Top to Bottom
                bDone = uaPts(lIdx - 1).y > lHeight And uaPts(lIdx - 2).y > lHeight
            Case iIncY < 0  'Moving Bottom to Top
                bDone = uaPts(lIdx - 1).y < 0 And uaPts(lIdx - 2).y < 0
        End Select
        If (lIdx Mod 1000) = 0 Then
            ReDim Preserve uaPts(UBound(uaPts) + 1000)
        End If
    Wend
    
    'Free excess memory (may have 1001 points dimmed to 2000)
    ReDim Preserve uaPts(lIdx - 1)
    
    'Create the array of colors blending from mvarColor1 to mvarColor2
    lRet = BlendColors(mvarColor1, mvarColor2, lIdx / 2, laColors)
    
    'Now draw each line in it's own color
    For lIdx = 0 To UBound(uaPts) - 1 Step 2
        'Move to next point
        lRet = MoveToEx(hDC, uaPts(lIdx).x, uaPts(lIdx).y, uTmpPt)
        'Create the colored pen and select it into the DC
        hPen = CreatePen(PS_SOLID, 1, laColors(Int(lIdx / 2)))
        hOldPen = SelectObject(hDC, hPen)
        'Draw the line
        lRet = LineTo(hDC, uaPts(lIdx + 1).x, uaPts(lIdx + 1).y)
        'Get the pen back out of the DC and destroy it
        lRet = SelectObject(hDC, hOldPen)
        lRet = DeleteObject(hPen)
    Next lIdx
    
    DrawGradient = lIdx
    
NormalExit:
    'Free the memory
    Erase laColors
    Erase uaPts
    Erase uaTmpPts
    Exit Function
    
LocalError:
    MsgBox Err.Description, vbExclamation, "SuperGradient.cls"
    DrawGradient = 0
    Resume 'NormalExit
    
End Function

Private Sub RotatePoint(uAxisPt As PointSng, uRotatePt As PointSng, fDegrees As Single)
    Dim fDX         As Single
    Dim fDY         As Single
    Dim fRadians    As Single

    fRadians = fDegrees * RADS
    fDX = uRotatePt.x - uAxisPt.x
    fDY = uRotatePt.y - uAxisPt.y
    uRotatePt.x = uAxisPt.x + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
    uRotatePt.y = uAxisPt.y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))
    
End Sub



Private Function Div(ByVal dNumer As Double, ByVal dDenom As Double) As Double
    
'Divides dNumer by dDenom if dDenom <> 0
'Eliminates 'Division By Zero' error.

    If dDenom <> 0 Then
        Div = dNumer / dDenom
    Else
        Div = 0
    End If

End Function

Public Property Let Color1(ByVal lData As Long)
    
Dim lIdx As Long

    mvarColor1 = lData
    If mvarColor1 < 0 Then
        lIdx = (mvarColor1 And Not &H80000000)
        If lIdx >= 0 And lIdx <= 24 Then
            mvarColor1 = GetSysColor(lIdx)
        End If
    End If
    
End Property


Public Property Get Color1() As Long
    Color1 = mvarColor1
End Property



Public Property Let Color2(ByVal lData As Long)

Dim lIdx As Long

    mvarColor2 = lData
    If mvarColor2 < 0 Then
        lIdx = (mvarColor2 And Not &H80000000)
        If lIdx >= 0 And lIdx <= 24 Then
            mvarColor2 = GetSysColor(lIdx)
        End If
    End If

End Property


Public Property Get Color2() As Long
    Color2 = mvarColor2
End Property



Public Property Let Angle(ByVal fData As Single)
    
'Angles are counter-clockwise and may be
'any Single value from 0 to 359.999999999.

' 135  90 45
'    \ | /
'180 --o-- 0
'    / | \
' 235 270 315

    'Correct angle to ensure between 0 and 359.999999999
    mvarAngle = CDbl(fData) - Int(Int(CDbl(fData) / 360#) * 360#)

End Property


Public Property Get Angle() As Single
    Angle = mvarAngle
End Property

Private Sub Class_Initialize()
    mvarColor1 = mvar_def_Color1
    mvarColor2 = mvar_def_Color2
    mvarAngle = mvar_def_Angle
End Sub


