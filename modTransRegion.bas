Attribute VB_Name = "modTransRegion"
'The code for the TransRegion.dll is modified from code by Chris Yates.  however, I will release it later on C/C++ of Planet
'Source Code.  However, in the meantime you may use the DLL with no warranty whatsoever and without any royalty fees

Declare Sub MakeTransparent Lib "TransRegion.dll" (ByVal WinHandle As Long, ByVal SrcHandle As Long, ByVal Red As Integer, ByVal Blue As Integer, ByVal Green As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal dat As Integer)

Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetParent Lib "User32" (ByVal hwnd As Long) As Long
Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long

'For moving the Mouse
Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long


'For the Window
'--------------
'Put the window at the bottom of the Z-order.
Public Const HWND_BOTTOM = 1
'Put the window below all topmost windows and above all non-topmost windows.
Public Const HWND_NOTOPMOST = -2
'Put the window at the top of the Z-order.
Public Const HWND_TOP = 0
'Make the window topmost (above all other windows) permanently.
Public Const HWND_TOPMOST = -1

'For the Flags
'--------------
'Same as SWP_FRAMECHANGED.
Public Const SWP_DRAWFRAME = &H20
'Fully redraw the window in its new position.
Public Const SWP_FRAMECHANGED = &H20
'Hide the window from the screen.
Public Const SWP_HIDEWINDOW = &H80
'Do not make the window active after moving it unless it was already the active window.
Public Const SWP_NOACTIVATE = &H10
'Do not redraw anything drawn on the window after it is moved.
Public Const SWP_NOCOPYBITS = &H100
'Do not move the window.
Public Const SWP_NOMOVE = &H2
'Do not resize the window.
Public Const SWP_NOSIZE = &H1
'Do not remove the image of the window in its former position, effectively leaving a ghost image on the screen.
Public Const SWP_NOREDRAW = &H8
'Do not change the window's position in the Z-order.
Public Const SWP_NOZORDER = &H4
'Show the window if it is hidden.
Public Const SWP_SHOWWINDOW = &H40


'Types
'--------------
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


