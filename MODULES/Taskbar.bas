Attribute VB_Name = "Taskbar"
'Option Explicit

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40
Const WS_CHILD = 1073741824
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const SW_HIDE = 0
Const SW_NORMAL = 1
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim tWnd As Long, bWnd As Long, ncWnd As Long
Dim rtn As Long

Sub StartBaru(ByVal txt)
    Dim R As RECT
    tWnd = FindWindow("Shell_TrayWnd", vbNullString)
    bWnd = FindWindowEx(tWnd, ByVal 0&, "BUTTON", vbNullString)
    GetWindowRect bWnd, R
    ncWnd = CreateWindowEx(ByVal 0&, "BUTTON", txt, WS_CHILD, 0, 0, R.Right - R.Left, R.Bottom - R.Top, tWnd, ByVal 0&, App.hInstance, ByVal 0&)
    ShowWindow ncWnd, SW_NORMAL
    ShowWindow bWnd, SW_HIDE
End Sub

Sub Normal()
    ShowWindow bWnd, SW_NORMAL
    DestroyWindow ncWnd
End Sub

Function SembunyiStart()
   OurParent& = FindWindow("Shell_TrayWnd", "")
   OurHandle& = FindWindowEx(OurParent&, 0, "Button", vbNullString)
   ShowWindow OurHandle&, 0
End Function

Function TampilkanStart()
   OurParent& = FindWindow("Shell_TrayWnd", "")
   OurHandle& = FindWindowEx(OurParent&, 0, "Button", vbNullString)
   ShowWindow OurHandle&, 5
End Function


Sub SembunyiTaksbar()
    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Sub

Sub TampilkanTaksbar()
    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Sub

