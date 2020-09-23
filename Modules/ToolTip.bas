Attribute VB_Name = "ToolTip"
Option Explicit

Public Const TOOLTIPS_CLASS = "tooltips_class32"

' Styles
Public Const TTS_ALWAYSTIP = &H1
Public Const TTS_NOPREFIX = &H2

Public Const TTM_SETMAXTIPWIDTH        As Long = (&H400 + 24)

''Windows API Constants
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1

Public Type TOOLINFO
  cbSize As Long
  uFlags As TT_Flags
  hwnd As Long
  uId As Long
  RECT As RECT
  hinst As Long
  lpszText As String   ' Long
#If (WIN32_IE >= &H300) Then
  lParam As Long
#End If
End Type   ' TOOLINFO

Private Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
    hIconSm As Long
End Type

Public Enum TT_Flags
  TTF_IDISHWND = &H1
  TTF_CENTERTIP = &H2
  TTF_RTLREADING = &H4
  TTF_SUBCLASS = &H10
#If (WIN32_IE >= &H300) Then
  TTF_TRACK = &H20
  TTF_ABSOLUTE = &H80
  TTF_TRANSPARENT = &H100
  TTF_DI_SETITEM = &H8000&        ' valid only on the TTN_NEEDTEXT callback
#End If     ' WIN32_IE >= =&H0300
End Enum   ' TT_Flags

Public Enum TT_Msgs
  TTM_ACTIVATE = (WM_USER + 1)
  TTM_SETDELAYTIME = (WM_USER + 3)
  TTM_RELAYEVENT = (WM_USER + 7)
  TTM_GETTOOLCOUNT = (WM_USER + 13)
  TTM_WINDOWFROMPOINT = (WM_USER + 16)
    
#If UNICODE Then
  TTM_ADDTOOL = (WM_USER + 50)
  TTM_DELTOOL = (WM_USER + 51)
  TTM_NEWTOOLRECT = (WM_USER + 52)
  TTM_GETTOOLINFO = (WM_USER + 53)
  TTM_SETTOOLINFO = (WM_USER + 54)
  TTM_HITTEST = (WM_USER + 55)
  TTM_GETTEXT = (WM_USER + 56)
  TTM_UPDATETIPTEXT = (WM_USER + 57)
  TTM_ENUMTOOLS = (WM_USER + 58)
  TTM_GETCURRENTTOOL = (WM_USER + 59)
#Else
  TTM_ADDTOOL = (WM_USER + 4)
  TTM_DELTOOL = (WM_USER + 5)
  TTM_NEWTOOLRECT = (WM_USER + 6)
  TTM_GETTOOLINFO = (WM_USER + 8)
  TTM_SETTOOLINFO = (WM_USER + 9)
  TTM_HITTEST = (WM_USER + 10)
  TTM_GETTEXT = (WM_USER + 11)
  TTM_UPDATETIPTEXT = (WM_USER + 12)
  TTM_ENUMTOOLS = (WM_USER + 14)
  TTM_GETCURRENTTOOL = (WM_USER + 15)
#End If   ' UNICODE

#If (WIN32_IE >= &H300) Then
  TTM_TRACKACTIVATE = (WM_USER + 17)       ' wParam = TRUE/FALSE start end  lparam = LPTOOLINFO
  TTM_TRACKPOSITION = (WM_USER + 18)       ' lParam = dwPos
  TTM_SETTIPBKCOLOR = (WM_USER + 19)
  TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
  TTM_GETDELAYTIME = (WM_USER + 21)
  TTM_GETTIPBKCOLOR = (WM_USER + 22)
  TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
  TTM_SETMAXTIPWIDTH = (WM_USER + 24)
  TTM_GETMAXTIPWIDTH = (WM_USER + 25)
  TTM_SETMARGIN = (WM_USER + 26)           ' lParam = lprc
  TTM_GETMARGIN = (WM_USER + 27)           ' lParam = lprc
  TTM_POP = (WM_USER + 28)
#End If   ' (WIN32_IE >= &H300)

#If (WIN32_IE >= &H400) Then
  TTM_UPDATE = (WM_USER + 29)
#End If
End Enum   ' TT_Msgs


Public Enum TT_DelayTime
  TTDT_AUTOMATIC = 0
  TTDT_RESHOW = 1
  TTDT_AUTOPOP = 2
  TTDT_INITIAL = 3
End Enum

Public Enum ttDelayTimeConstants
  ttDelayDefault = TTDT_AUTOMATIC '= 0
  ttDelayInitial = TTDT_INITIAL '= 3
  ttDelayShow = TTDT_AUTOPOP '= 2
  ttDelayReshow = TTDT_RESHOW '= 1
  ttDelayMask = 3
End Enum

'-----------Window Style Constants
Public Const WS_BORDER = &H800000
Public Const WS_CHILD = &H40000000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_VISIBLE = &H10000000
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_VSCROLL = &H200000
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_WINDOWEDGE = &H100&
Public Const WS_SIZEBOX = &H40000
Public Const WS_EX_DLGMODALFRAME = &H1&

Private Const CS_HREDRAW = &H2
Private Const CS_VREDRAW = &H1
Private Const CS_PARENTDC = &H80
Private Const WS_EX_APPWINDOW = &H40000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const IDC_ARROW = &H7F00
Private Const COLOR_WINDOW = &H5
Private Const SW_SHOW = &H5
Private Const WM_DESTROY = &H2
Private Const WM_PAINT = &HF
Private Const WM_CREATE = &H1
Private Const DT_CENTER = &H1
Private Const SS_SUNKEN = &H1000
Private Const SS_CENTER = &H1
Private Const LBS_NOTIFY = &H1
Private Const LB_ADDSTRING = &H180
Private Const WM_COMMAND = &H111
Private Const CBS_DROPDOWNLIST = &H3
Private Const CBS_AUTOHSCROLL = &H40
Private Const CBS_HASSTRINGS = &H200
Private Const CB_ADDSTRING = &H143
Private Const CBS_DISABLENOSCROLL = &H800&
Private Const CB_SETCURSEL = &H14E
Private Const LBN_SELCHANGE = &H1
Private Const LB_GETTEXT = &H189
Private Const LB_GETCURSEL = &H188
Private Const WM_TIMER = &H113

' dims for the dynamic controls
Dim hWndButton As Long
Private Const IDC_BUTTON = &H1000
Dim hWndEditBox As Long
Private Const IDC_EDIT = &H1001
Dim hWndStatic As Long
Private Const IDC_STATIC = &H1002
Dim hWndList As Long
Private Const IDC_LIST = &H1003
Dim hWndCombo As Long
Private Const IDC_COMBO = &H1004
Dim hWndStaticTimer As Long
Private Const IDC_STATIC_TIMER = &H1005
Private Const ID_TIMER = &HCAFEBABE


Public Function CreateToolTip(lText As String, hwnd As Long) As Boolean
    Dim m_hWndTooltip As Long
    m_hWndTooltip = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, vbNullString, WS_POPUP Or TTS_NOPREFIX Or TTS_ALWAYSTIP, 0, 0, 100, 100, hwnd, 0, App.hInstance, ByVal 0)
    '--- make tooltips multi-line
    'Call SendMessage(m_hWndTooltip, TTM_SETMAXTIPWIDTH, 0&, ByVal &H7FFF&)
    'm_hWndTooltip = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, vbNullString, WS_POPUP Or TTS_NOPREFIX, 0, 0, 100, 100, hwnd, 0, App.hInstance, ByVal 0)
    m_hWndTooltip = CreateWindowEx(WS_EX_TOPMOST, _
                              TOOLTIPS_CLASS, _
                              vbNullString, _
                              WS_POPUP Or TTS_NOPREFIX Or TTS_ALWAYSTIP, _
                              100, 100, 200, 200, 0, 0, App.hInstance, 0)
    UpdateWindow m_hWndTooltip
    ShowWindow m_hWndTooltip, SW_SHOW
End Function

