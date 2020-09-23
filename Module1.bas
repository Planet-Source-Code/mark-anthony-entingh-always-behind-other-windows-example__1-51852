Attribute VB_Name = "Module1"
'alwaysontop
    Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
    Public Const HWND_NOTOPMOST = -2
    Public Const HWND_TOPMOST = -1
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
'close an app code
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Const WM_CLOSE = &H10



Public Sub AlwaysOnBottom(frm As Form)
    On Error Resume Next
    hwnd1 = FindWindow("SysListView32", vbNullString)
    cx = frm.ScaleWidth
    cy = frm.ScaleHeight
    X = frm.Left
    Y = frm.Top
    Call SetWindowPos(frm.hwnd, 1, X, Y, 360, 382, SWP_NOMOVE)
End Sub
