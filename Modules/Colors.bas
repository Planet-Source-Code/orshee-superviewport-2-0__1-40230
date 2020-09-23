Attribute VB_Name = "Colors"
Option Explicit

Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, _
                                                               ByVal lHPalette As Long, _
                                                               lColorRef As Long) As Long

Public Enum stcColorIndex
    ciRed = 0
    ciGreen = 1
    ciBlue = 2
End Enum
    

Private Function ConvertHexCharacters(sChar As String) As Integer
    Select Case sChar
        Case "0" To "9"
            ConvertHexCharacters = Val(sChar)
        Case "A" To "F"
            ConvertHexCharacters = Asc(sChar) - 55
        Case Else
            Exit Function
    End Select
End Function

Public Function GetRGBColor(ByVal LongColorValue As Long, ColorIndex As stcColorIndex) As Long
    Dim myRedValue As String
    Dim myGreenValue As String
    Dim myBlueValue As String
    Dim myHexString1 As String
    Dim myHexString2 As String
    Dim myString As String * 6
    
    myHexString2 = Hex$(LongColorValue)
    myString = "000000"
    
    RSet myString = myHexString2
    
    myRedValue = 16 * ConvertHexCharacters(Mid$(myString, 5, 1)) + ConvertHexCharacters(Mid$(myString, 6, 1))
    myGreenValue = 16 * ConvertHexCharacters(Mid$(myString, 3, 1)) + ConvertHexCharacters(Mid$(myString, 4, 1))
    myBlueValue = 16 * ConvertHexCharacters(Mid$(myString, 1, 1)) + ConvertHexCharacters(Mid$(myString, 2, 1))
    
    Select Case ColorIndex
    Case ciRed
        GetRGBColor = myRedValue
    Case ciGreen
        GetRGBColor = myGreenValue
    Case ciBlue
        GetRGBColor = myBlueValue
    End Select
    'GetRGBColor = Str$(myRedValue) & Str$(myGreenValue) & Str$(myBlueValue)
End Function

Public Function TranslateOLE_Color(ByVal Value As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(Value, hPal, TranslateOLE_Color) Then
        TranslateOLE_Color = -1
    End If
End Function


