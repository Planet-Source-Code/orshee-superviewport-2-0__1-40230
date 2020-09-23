VERSION 5.00
Begin VB.UserControl SuperViewPort 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1425
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   ScaleHeight     =   76
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   95
   ToolboxBitmap   =   "SuperViewPort.ctx":0000
End
Attribute VB_Name = "SuperViewPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements iSuperSubClasser


'***[Enumerations]*********************************************************************
Public Enum enuBorderStyle
    bsNoBorder = 0
    bsFixedSingle = 1
End Enum

Public Enum enuScrollBars
    sbNone = 0
    sbHorizontal = 1
    sbVertical = 2
    sbBoth = 3
End Enum


'***[Events]*********************************************************************
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Move()
Event Resize()
Event ScrollH(Value As Long)
Event ScrollV(Value As Long)
Event MouseScroll(CurrentH As Long, CurrentV As Long)
Event EndScroll()


'***[Constants]*********************************************************************
Const HSCROLLBAR As Long = 0
Const VSCROLLBAR As Long = 1


'***[Default Constants]*********************************************************************
Const mvar_def_BackColor            As Long = vb3DShadow '&H80000010
Const mvar_def_ForeColor            As Long = vb3DFace '&H8000000F
Const mvar_def_ScrollBars           As Long = sbNone
Const mvar_def_ViewPortWidth        As Long = 1000
Const mvar_def_ViewPortHeight       As Long = 1000
Const mvar_def_SmallChangeH         As Long = 10
Const mvar_def_SmallChangeV         As Long = 10
Const mvar_def_LargeChangeH         As Long = 100
Const mvar_def_LargeChangeV         As Long = 100
Const mvar_def_GradientEnabled      As Boolean = False
Const mvar_def_GradientAngle        As Long = 270
Const mvar_def_MouseTrack           As Boolean = False

'***[Shared Variables]*********************************************************************
Private mvarWindowSubClasser        As SuperSubClasser
Private mvarScrollBars              As enuScrollBars
Private mvarViewPortWidth           As Long
Private mvarViewPortHeight          As Long
Private mvarCurrentPosH             As Long
Private mvarCurrentPosV             As Long
Private mvarSmallChangeH            As Long
Private mvarSmallChangeV            As Long
Private mvarLargeChangeH            As Long
Private mvarLargeChangeV            As Long
Private mvarScrollInfo              As SCROLLINFO

Private mvarGradient                As SuperGradient
Private mvarGradientEnabled         As Boolean
Private mvarGradientAngle           As Long

Private mvarMouseTrack              As Boolean


'***[Storage Variables]*********************************************************************
Private mvarHScrollAbs As Long
Private mvarVScrollAbs As Long


'***[Life Control]*********************************************************************
Private Sub iSuperSubClasser_Before(lHandled As Long, lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    'DoNothing
End Sub

Private Sub iSuperSubClasser_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    Select Case uMsg
        Case WM_MOVE
            RaiseEvent Move
        Case WM_SIZE
            RaiseEvent Resize
        Case WM_HSCROLL
            With Me
                Select Case LOWORD(wParam)
                    Case SB_BOTTOM
                        .CurrentPosH = .ViewPortWidth
                    Case SB_TOP
                        .CurrentPosH = 0
                    Case SB_LINEDOWN
                        If mvarCurrentPosH < mvarViewPortWidth Then
                            .CurrentPosH = .CurrentPosH + mvarSmallChangeH
                        Else
                            .CurrentPosH = mvarViewPortWidth
                        End If
                    Case SB_LINEUP
                        If mvarCurrentPosH > 0 Then
                            .CurrentPosH = .CurrentPosH - mvarSmallChangeH
                        Else
                            .CurrentPosH = 0
                        End If
                    Case SB_PAGEDOWN
                        If mvarCurrentPosH < mvarViewPortWidth Then
                            .CurrentPosH = .CurrentPosH + mvarLargeChangeH
                        Else
                            .CurrentPosH = mvarViewPortWidth
                        End If
                    Case SB_PAGEUP
                        If mvarCurrentPosH > mvarLargeChangeH Then
                            .CurrentPosH = .CurrentPosH - mvarLargeChangeH
                        Else
                            .CurrentPosH = 0
                        End If
                    Case SB_THUMBPOSITION, SB_THUMBTRACK
                        .CurrentPosH = HIWORD(wParam)
                        'CreateToolTip CStr(mvarCurrentPosH), UserControl.hwnd
                   Case SB_ENDSCROLL
                        'Each time any scroll method ends this message is passed
                        'We use it for EndScroll event that might be usefull
                        RaiseEvent EndScroll
                End Select
            End With
            
            lReturn = HTHSCROLL
        Case WM_VSCROLL
            With Me
                Select Case LOWORD(wParam)
                    Case SB_BOTTOM
                        .CurrentPosV = .ViewPortHeight
                    Case SB_TOP
                        .CurrentPosV = 0
                    Case SB_LINEDOWN
                        If mvarCurrentPosV < mvarViewPortHeight Then
                            .CurrentPosV = .CurrentPosV + mvarSmallChangeV
                        Else
                            CurrentPosV = mvarViewPortHeight
                        End If
                    Case SB_LINEUP
                        If mvarCurrentPosV > 0 Then
                            .CurrentPosV = .CurrentPosV - mvarSmallChangeV
                        Else
                            CurrentPosV = 0
                        End If
                    Case SB_PAGEDOWN
                        If mvarCurrentPosV < mvarViewPortHeight Then
                            .CurrentPosV = .CurrentPosV + mvarLargeChangeV
                        Else
                            CurrentPosV = mvarViewPortHeight
                        End If
                    Case SB_PAGEUP
                        If mvarCurrentPosV > mvarLargeChangeV Then
                            .CurrentPosV = .CurrentPosV - mvarLargeChangeV
                        Else
                            CurrentPosV = 0
                        End If
                    Case SB_THUMBPOSITION, SB_THUMBTRACK
                        .CurrentPosV = HIWORD(wParam)
                    Case SB_ENDSCROLL
                        'Each time any scroll method ends this message is passed
                        'We use it for EndScroll event that might be usefull
                        RaiseEvent EndScroll
                End Select
            End With
'        Case WM_NCMOUSEMOVE
'            Select Case LOWORD(wParam)
'            Case 6
'                'mvarToolTip.TipText = CStr(mvarCurrentPosH)
'                'Set mvarToolTip.ParentControl = UserControl.Extender
'                CreateToolTip CStr(mvarCurrentPosH), UserControl.hWnd
'                'mvarToolTip.Create
'            Case 7
'                'mvarToolTip.TipText = CStr(mvarCurrentPosV)
'                'Set mvarToolTip.ParentControl = UserControl.Extender
'                'mvarToolTip.Create
'                CreateToolTip CStr(mvarCurrentPosH), UserControl.hWnd
'
'            End Select
            lReturn = HTVSCROLL
    '    Case Else
    '        Call DefWindowProc(hwnd, uMsg, 0, 0)
    End Select
End Sub

Private Sub UserControl_Initialize()
    Set mvarWindowSubClasser = New SuperSubClasser
    Set mvarGradient = New SuperGradient
    
    'Add only messages needed for processing elemental events
    'above which others are made
    With mvarWindowSubClasser
        .AddMsg (WM_MOVE)
        .AddMsg (WM_SIZE)
        .AddMsg (WM_HSCROLL)
        .AddMsg (WM_VSCROLL)
        '.AddMsg (WM_NCMOUSEMOVE)
        .Subclass UserControl.hwnd, Me, False
    End With
        
    RenderControl
End Sub

Private Sub UserControl_Terminate()
    Set mvarWindowSubClasser = Nothing
    Set mvarGradient = Nothing
End Sub


Private Sub UserControl_Show()
    DrawGradient
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim myCtl As Control
    
    RaiseEvent MouseMove(Button, Shift, x, y)
    
    If MouseTrack = True And Button = vbLeftButton And mvarScrollBars <> sbNone Then
        'Enable mouse draging and set appropriate cursor
        'Mouse cursor will be reset in MouseUp event
        Select Case mvarScrollBars
        Case sbHorizontal
            Screen.MousePointer = vbSizeWE
        Case sbVertical
            Screen.MousePointer = vbSizeNS
        Case sbBoth
            Screen.MousePointer = vbSizeAll
        End Select
        
        CurrentPosH = x
        CurrentPosV = y
        SetControlsPosition
        RaiseEvent MouseScroll(mvarCurrentPosH, mvarCurrentPosV)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Screen.MousePointer <> vbDefault Then
        Screen.MousePointer = vbDefault
        RaiseEvent EndScroll
    End If
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub


Private Sub SetControlsPosition()
    Dim myCtl As Control
    
    On Local Error Resume Next
    'Set contained controls new position
    For Each myCtl In ContainedControls
        myCtl.Left = myCtl.Left + (mvarHScrollAbs - mvarCurrentPosH) * Screen.TwipsPerPixelX
        myCtl.Top = myCtl.Top + (mvarVScrollAbs - mvarCurrentPosV) * Screen.TwipsPerPixelY
    Next
    mvarHScrollAbs = mvarCurrentPosH
    mvarVScrollAbs = mvarCurrentPosV

End Sub

Private Sub UserControl_Resize()
    DrawGradient
    RenderControl
End Sub

Private Sub UserControl_InitProperties()
    'Basic UserControl initialisation event in which we set
    'all properties to default values
    UserControl.ForeColor = mvar_def_ForeColor
    UserControl.BackColor = mvar_def_BackColor
    
    mvarScrollBars = mvar_def_ScrollBars
    mvarViewPortWidth = mvar_def_ViewPortWidth
    mvarViewPortHeight = mvar_def_ViewPortHeight

    mvarSmallChangeH = mvar_def_SmallChangeH
    mvarSmallChangeV = mvar_def_SmallChangeV
    LargeChangeH = mvar_def_LargeChangeH
    LargeChangeV = mvar_def_LargeChangeV

    mvarGradientEnabled = mvar_def_GradientEnabled
    mvarGradientAngle = mvar_def_GradientAngle
    
    mvarMouseTrack = mvar_def_MouseTrack
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", bsFixedSingle)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", mvar_def_ForeColor)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", mvar_def_BackColor)
    
    ScrollBars = PropBag.ReadProperty("ScrollBars", mvar_def_ScrollBars)
    
    ViewPortWidth = PropBag.ReadProperty("ViewPortWidth", mvar_def_ViewPortWidth)
    ViewPortHeight = PropBag.ReadProperty("ViewPortHeight", mvar_def_ViewPortHeight)
    
    mvarSmallChangeH = PropBag.ReadProperty("SmallChangeH", mvar_def_SmallChangeH)
    mvarSmallChangeV = PropBag.ReadProperty("SmallChangeV", mvar_def_SmallChangeV)
    LargeChangeH = PropBag.ReadProperty("LargeChangeH", mvar_def_LargeChangeH)
    LargeChangeV = PropBag.ReadProperty("LargeChangeV", mvar_def_LargeChangeV)
    
    GradientEnabled = PropBag.ReadProperty("GradientEnabled", mvar_def_GradientEnabled)
    GradientAngle = PropBag.ReadProperty("GradientAngle", mvar_def_GradientAngle)
    
    mvarMouseTrack = PropBag.ReadProperty("MouseTrack", mvar_def_MouseTrack)
    
    RenderControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BorderStyle", UserControl.BorderStyle, bsFixedSingle
    PropBag.WriteProperty "ForeColor", UserControl.ForeColor, mvar_def_ForeColor
    PropBag.WriteProperty "BackColor", UserControl.BackColor, mvar_def_BackColor
    
    PropBag.WriteProperty "ScrollBars", mvarScrollBars, mvar_def_ScrollBars
    
    PropBag.WriteProperty "ViewPortWidth", mvarViewPortWidth, mvar_def_ViewPortWidth
    PropBag.WriteProperty "ViewPortHeight", mvarViewPortHeight, mvar_def_ViewPortHeight

    PropBag.WriteProperty "SmallChangeH", mvarSmallChangeH, mvar_def_SmallChangeH
    PropBag.WriteProperty "SmallChangeV", mvarSmallChangeV, mvar_def_SmallChangeV
    PropBag.WriteProperty "LargeChangeH", mvarLargeChangeH, mvar_def_LargeChangeH
    PropBag.WriteProperty "LargeChangeV", mvarLargeChangeV, mvar_def_LargeChangeV

    PropBag.WriteProperty "GradientEnabled", mvarGradientEnabled, mvar_def_GradientEnabled
    PropBag.WriteProperty "GradientAngle", mvarGradientAngle, mvar_def_GradientAngle

    PropBag.WriteProperty "MouseTrack", mvarMouseTrack, mvar_def_MouseTrack
End Sub

Private Sub RenderControl()
    Dim myhWnd As Long
    
    myhWnd = UserControl.hwnd
    
    Select Case mvarScrollBars
        Case sbNone
            ShowScrollBar myhWnd, HSCROLLBAR, API_FALSE
            ShowScrollBar myhWnd, VSCROLLBAR, API_FALSE
        Case sbHorizontal
            ShowScrollBar myhWnd, HSCROLLBAR, API_TRUE
            ShowScrollBar myhWnd, VSCROLLBAR, API_FALSE
        Case sbVertical
            ShowScrollBar myhWnd, VSCROLLBAR, API_TRUE
            ShowScrollBar myhWnd, HSCROLLBAR, API_FALSE
        Case sbBoth
            ShowScrollBar myhWnd, HSCROLLBAR, API_TRUE
            ShowScrollBar myhWnd, VSCROLLBAR, API_TRUE
    End Select
End Sub

Private Sub DrawGradient()
    Dim myhWnd As Long
    Dim myhDC As Long
    
    myhWnd = UserControl.hwnd
    myhDC = UserControl.hDC
    
    If mvarGradientEnabled = True Then
        With mvarGradient
            .Angle = mvarGradientAngle
            .Color1 = ForeColor
            .Color2 = BackColor
            .Draw myhWnd, myhDC
        End With
        UserControl.Refresh
    Else
        Cls
    End If

End Sub

'***[Properties]*********************************************************************
Public Property Let ScrollBars(Value As enuScrollBars)
    mvarScrollBars = Value
    RenderControl
    DrawGradient
End Property

Public Property Get ScrollBars() As enuScrollBars
    ScrollBars = mvarScrollBars
End Property

Public Property Let ViewPortWidth(Value As Long)
    mvarViewPortWidth = Value
    
    With mvarScrollInfo
        .fMask = SIF_RANGE
        .nMin = 0
        .nMax = Value '+ mvarLargeChangeH
    End With
    
    SetScrollInfo hwnd, SB_HORZ, mvarScrollInfo, True
    
    RenderControl
End Property

Public Property Get ViewPortWidth() As Long
    ViewPortWidth = mvarViewPortWidth
End Property

Public Property Let ViewPortHeight(Value As Long)
    mvarViewPortHeight = Value
    
    With mvarScrollInfo
        .fMask = SIF_RANGE
        .nMin = 0
        .nMax = Value '+ mvarLargeChangeV
    End With
    
    SetScrollInfo hwnd, SB_VERT, mvarScrollInfo, True
    
    RenderControl
End Property

Public Property Get ViewPortHeight() As Long
    ViewPortHeight = mvarViewPortHeight
End Property

Public Property Let CurrentPosH(Value As Long)
    'First check which scrollbar is visible and is it posssble to set the position
    If mvarScrollBars = sbNone Or mvarScrollBars = sbVertical Then
        Exit Property
    End If
    
    If Value > mvarViewPortWidth - mvarLargeChangeH Then
        mvarCurrentPosH = mvarViewPortWidth - mvarLargeChangeH
    Else
        mvarCurrentPosH = Value
    End If
 
    If Value < 0 Then mvarCurrentPosH = 0
    
    With mvarScrollInfo
        .fMask = SIF_POS
        .nPos = Value
    End With
    
    SetScrollInfo hwnd, SB_HORZ, mvarScrollInfo, True

    SetControlsPosition
    
    RaiseEvent ScrollH(mvarCurrentPosH)
End Property

Public Property Get CurrentPosH() As Long
    CurrentPosH = mvarCurrentPosH
End Property

Public Property Let CurrentPosV(Value As Long)
    'First check which scrollbar is visible and is it posssble to set the position
    If mvarScrollBars = sbNone Or mvarScrollBars = sbHorizontal Then
        Exit Property
    End If
    
    If Value > mvarViewPortHeight - mvarLargeChangeV Then
        mvarCurrentPosV = mvarViewPortHeight - mvarLargeChangeV
    Else
        mvarCurrentPosV = Value
    End If

    If Value < 0 Then mvarCurrentPosV = 0

    With mvarScrollInfo
        .fMask = SIF_POS
        .nPos = Value
    End With
    
    SetScrollInfo hwnd, SB_VERT, mvarScrollInfo, True
    
    SetControlsPosition
    
    RaiseEvent ScrollV(mvarCurrentPosV)
End Property

Public Property Get CurrentPosV() As Long
    CurrentPosV = mvarCurrentPosV
End Property

'***
Public Property Let SmallChangeH(Value As Long)
    mvarSmallChangeH = Value
    RenderControl
End Property

Public Property Get SmallChangeH() As Long
    SmallChangeH = mvarSmallChangeH
End Property

Public Property Let SmallChangeV(Value As Long)
    mvarSmallChangeV = Value
    RenderControl
End Property

Public Property Get SmallChangeV() As Long
    SmallChangeV = mvarSmallChangeV
End Property

Public Property Let LargeChangeH(Value As Long)
    If Value > mvarViewPortWidth Then
        mvarLargeChangeH = mvarViewPortWidth
    Else
        mvarLargeChangeH = Value
    End If

    With mvarScrollInfo
        .fMask = SIF_PAGE
        .nPage = Value
    End With
    
    RenderControl
    SetScrollInfo hwnd, SB_HORZ, mvarScrollInfo, True
End Property

Public Property Get LargeChangeH() As Long
    LargeChangeH = mvarLargeChangeH
End Property

Public Property Let LargeChangeV(Value As Long)
    If Value > mvarViewPortHeight Then
        mvarLargeChangeV = mvarViewPortHeight
    Else
        mvarLargeChangeV = Value
    End If

    With mvarScrollInfo
        .fMask = SIF_PAGE
        .nPage = Value
    End With
    
    RenderControl
    SetScrollInfo hwnd, SB_VERT, mvarScrollInfo, True
End Property

Public Property Get LargeChangeV() As Long
    LargeChangeV = mvarLargeChangeV
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
    UserControl.ForeColor() = Value
    DrawGradient
    PropertyChanged "ForeColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    UserControl.BackColor() = Value
    DrawGradient
    PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As enuBorderStyle
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal Value As enuBorderStyle)
    UserControl.BorderStyle() = Value
    DrawGradient
    PropertyChanged "BorderStyle"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    UserControl.Enabled() = Value
    PropertyChanged "Enabled"
End Property

Public Property Get GradientEnabled() As Boolean
    GradientEnabled = mvarGradientEnabled
End Property

Public Property Let GradientEnabled(ByVal Value As Boolean)
    mvarGradientEnabled = Value
    PropertyChanged "GradientEnabled"
    DrawGradient
End Property

Public Property Let GradientAngle(Value As Long)
    mvarGradientAngle = Value
    PropertyChanged "GradientAngle"
    DrawGradient
End Property

Public Property Get GradientAngle() As Long
    GradientAngle = mvarGradientAngle
End Property

Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550
    RenderControl
End Sub

Public Property Get MouseTrack() As Boolean
    MouseTrack = mvarMouseTrack
End Property

Public Property Let MouseTrack(ByVal Value As Boolean)
    mvarMouseTrack = Value
    PropertyChanged "MouseTrack"
End Property

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Sub About()
Attribute About.VB_UserMemId = -552
    MsgBox "Advanced Controls - Super ViewPort 2.0", vbInformation, App.Title
End Sub



