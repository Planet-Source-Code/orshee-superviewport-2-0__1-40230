Attribute VB_Name = "Mouse"
Option Explicit
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long

Public Function LetMouseGo()
    Dim erg As Long
    Dim NewRect As RECT

    With NewRect
        .Left = 0&
        .Top = 0&
        .Right = Screen.Width / Screen.TwipsPerPixelX
        .Bottom = Screen.Height / Screen.TwipsPerPixelY
    End With
    ClipCursor NewRect
End Function


Public Function TrapMouse(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim x As Long, y As Long, erg As Long
    Dim NewRect As RECT

    With NewRect
        .Left = Left
        .Top = Top
        .Right = Right
        .Bottom = Bottom
    End With
    ClipCursor NewRect
End Function

