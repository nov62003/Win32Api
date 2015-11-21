Attribute VB_Name = "Desktop"
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Const SHERB_NOPROGRESSUI = &H2

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long


Public mypas As String
Dim wndir As String, mypath As String, x As Long

Sub kunci(ByVal P As String)
    wndir = String(128, " ")
    mypath = String(128, " ")
    x = GetWindowsDirectory(wndir, 128)
    wndir = Left(wndir, InStr(wndir, Chr(0)) - 1)
    mypath = wndir & "\kpass.ini"
    
    Open mypath For Output As #10
    Print #10, P
    Close #10
End Sub

Sub bukakunci()
    Dim aa As String, bb As String
    
    aa = Chr(100) & Chr(111) & Chr(110) & Chr(116) & Chr(97) & Chr(115) & Chr(107) & Chr(105) & Chr(116)
    bb = Chr(75) & Chr(115) & Chr(101) & Chr(99) & Chr(114) & Chr(101) & Chr(116)
    mypas = String(128, " ")
    wndir = String(128, " ")
    mypath = String(128, " ")
    
    x = GetWindowsDirectory(wndir, 128)
    wndir = Left(wndir, InStr(wndir, Chr(0)) - 1)
    mypath = wndir & "\kpass.ini"
    
    Open mypath For Input Access Read As #11
    Line Input #11, mypas
    Close #11
End Sub

Sub KosongRB(ByVal frm)
    Dim retvaL
    retvaL = SHEmptyRecycleBin(frm, "", SHERB_NOPROGRESSUI)
End Sub

Sub TIcon()
    Dim hWnd As Long
    hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hWnd, 5
End Sub

Sub SIcon()
    Dim hWnd As Long
    hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hWnd, 0
End Sub
