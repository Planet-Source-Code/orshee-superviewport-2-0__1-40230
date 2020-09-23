Attribute VB_Name = "Globals"
Option Explicit

Public Const API_TRUE As Long = 1&
Public Const API_FALSE As Long = 0&

Public Const GWL_STYLE = (-16)
Public Const GWL_WNDPROC = (-4)

Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_SYSMENU = &H80000
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_OVERLAPPED = &H0&
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)

'Windows messages that we're going to filter for callback.
Public Const WM_NULL             As Long = &H0
Public Const WM_CREATE           As Long = &H1
Public Const WM_DESTROY          As Long = &H2
Public Const WM_MOVE             As Long = &H3
Public Const WM_SIZE             As Long = &H5
Public Const WM_ACTIVATE         As Long = &H6
Public Const WM_NCMOUSEMOVE      As Long = &HA0
Public Const WM_NCLBUTTONDOWN    As Long = &HA1
Public Const WM_NCLBUTTONUP      As Long = &HA2
Public Const WM_NCLBUTTONDBLCLK  As Long = &HA3
Public Const WM_NCRBUTTONDOWN    As Long = &HA4
Public Const WM_NCMBUTTONDOWN    As Long = &HA7
Public Const WM_NCMBUTTONUP      As Long = &HA8
Public Const WM_NCMBUTTONDBLCLK  As Long = &HA9
Public Const WM_NCHITTEST        As Long = &H84
Public Const WM_SYSCOMMAND       As Long = &H112
Public Const WM_MOUSEMOVE        As Long = &H200
Public Const WM_LBUTTONDOWN      As Long = &H201
Public Const WM_LBUTTONUP        As Long = &H202
Public Const WM_LBUTTONDBLCLK    As Long = &H203
Public Const WM_RBUTTONDOWN      As Long = &H204
Public Const WM_RBUTTONUP        As Long = &H205
Public Const WM_RBUTTONDBLCLK    As Long = &H206
Public Const WM_MBUTTONDOWN      As Long = &H207
Public Const WM_MBUTTONUP        As Long = &H208
Public Const WM_MBUTTONDBLCLK    As Long = &H209
Public Const WM_MOUSEWHEEL       As Long = &H20A
Public Const WM_PAINT            As Long = &HF
Public Const WM_USER             As Long = &H400

Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115

'Message parameters
Public Const HTBORDER = 18
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTCAPTION = 2
Public Const HTCLIENT = 1
Public Const HTERROR = (-2)
Public Const HTGROWBOX = 4
Public Const HTHSCROLL = 6
Public Const HTLEFT = 10
Public Const HTMAXBUTTON = 9
Public Const HTMENU = 5
Public Const HTMINBUTTON = 8
Public Const HTNOWHERE = 0
Public Const HTREDUCE = HTMINBUTTON
Public Const HTRIGHT = 11
Public Const HTSIZE = HTGROWBOX
Public Const HTSIZEFIRST = HTLEFT
Public Const HTSIZELAST = HTBOTTOMRIGHT
Public Const HTSYSMENU = 3
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTTRANSPARENT = (-1)
Public Const HTVSCROLL = 7
Public Const HTZOOM = HTMAXBUTTON

'ScrollBar flags
Public Const ESB_ENABLE_BOTH = &H0
Public Const ESB_DISABLE_BOTH = &H3
Public Const ESB_DISABLE_LEFT = &H1
Public Const ESB_DISABLE_RIGHT = &H2
Public Const ESB_DISABLE_UP = &H1
Public Const ESB_DISABLE_DOWN = &H2
Public Const ESB_DISABLE_LTUP = ESB_DISABLE_LEFT
Public Const ESB_DISABLE_RTDN = ESB_DISABLE_RIGHT

'ScrollBar Constants
Public Const SB_HORZ = 0
Public Const SB_VERT = 1
Public Const SB_CTL = 2

'ScrollBar Commands
Public Const SB_LINEUP As Long = 0&
Public Const SB_LINELEFT As Long = 0&
Public Const SB_LINEDOWN As Long = 1&
Public Const SB_LINERIGHT As Long = 1&
Public Const SB_PAGEUP As Long = 2&
Public Const SB_PAGELEFT As Long = 2&
Public Const SB_PAGEDOWN As Long = 3&
Public Const SB_PAGERIGHT As Long = 3&
Public Const SB_THUMBPOSITION As Long = 4&
Public Const SB_THUMBTRACK As Long = 5&
Public Const SB_TOP As Long = 6&
Public Const SB_LEFT As Long = 6&
Public Const SB_BOTTOM As Long = 7&
Public Const SB_RIGHT As Long = 7&
Public Const SB_ENDSCROLL As Long = 8&

Public Const SIF_RANGE = &H1
Public Const SIF_PAGE = &H2
Public Const SIF_POS = &H4
Public Const SIF_DISABLENOSCROLL = &H8
Public Const SIF_TRACKPOS = &H10
Public Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Public Const SW_SCROLLCHILDREN As Long = &H1
Public Const SW_INVALIDATE     As Long = &H2
Public Const SW_ERASE          As Long = &H4

Public Const PS_SOLID As Long = 0  'Solid Pen Style (Used for CreatePen())

Public Type POINT
    x As Long
    y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type


' Scroll Bar APIs
Public Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Public Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Public Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
'Following Scrollbar APIs are not used any more, and are available because of
'backward compatibility
'Public Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
'Public Declare Function SetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
'Public Declare Function GetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long) As Long
'Public Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
'Public Declare Function EnableScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long

'Window APIs
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function ScrollWindow Lib "user32" (ByVal hwnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As RECT, lpClipRect As RECT) As Long
Public Declare Function ScrollWindowEx Lib "user32" (ByVal hwnd As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As Any, ByVal fuScroll As Long) As Long
Public Declare Function ScrollDC Lib "user32" (ByVal hDC As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)

'Graphical APIs
Public Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINT) As Long
Public Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Public Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Ellipse Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Public Function LOWORD(ByVal Value As Long) As Integer
' Returns the low 16-bit integer from a 32-bit long integer
  CopyMemory LOWORD, Value, 2&
End Function

Public Function HIWORD(ByVal Value As Long) As Integer
' Returns the high 16-bit integer from a 32-bit long integer
  CopyMemory HIWORD, ByVal VarPtr(Value) + 2, 2&
End Function

