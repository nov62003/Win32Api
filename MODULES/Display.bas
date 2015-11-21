Attribute VB_Name = "Display"
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpInitData As DEVMODE, ByVal dwFlags As Long) As Long
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (lpszDeviceName As Any, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32

Public Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const DM_DISPLAYFLAGS = &H200000
Const DM_DISPLAYFREQUENCY = &H400000

Const BITSPIXEL = 12

' Keadaan untuk mengatur perubahan display
Const CDS_UPDATEREGISTRY = &H1
Const CDS_TEST = &H2
Const CDS_FULLSCREEN = &H4
Const CDS_GLOBAL = &H8
Const CDS_SET_PRIMARY = &H10
Const CDS_RESET = &H40000000
Const CDS_SETRECT = &H20000000
Const CDS_NORESET = &H10000000

' Kembalikan Nilai untuk mengatur perubahan display
Const DISP_CHANGE_SUCCESSFUL = 0
Const DISP_CHANGE_RESTART = 1
Const DISP_CHANGE_FAILED = -1
Const DISP_CHANGE_BADMODE = -2
Const DISP_CHANGE_NOTUPDATED = -3
Const DISP_CHANGE_BADFLAGS = -4
Const DISP_CHANGE_BADPARAM = -5

Const EWX_LogOff = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4

Public d() As DEVMODE

'Background
Dim ret As Long

Const SPIF_SENDWININICHANGE = &H2
Const SPIF_UPDATEINIFILE = &H1
Const SPIF_SETDESKWALLPAPER = 20

'screen saver
Const SPI_SETSCREENSAVEACTIVE = 17

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long


Sub AturResolusi(ByVal x As Long)
    Dim L As Long, flags As Long
    d(x).dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT
    flags = CDS_UPDATEREGISTRY
    L = ChangeDisplaySettings(d(x), flags)
    Select Case L
        Case DISP_CHANGE_RESTART
            L = MsgBox("Perubahan ini tidak akan berpengaruh sampai anda me-restart sistem.  Reboot sekarang?", vbYesNo)
            If L = vbYes Then
                flags = 0
                L = ExitWindowsEx(EWX_REBOOT, flags)
            End If
        Case DISP_CHANGE_SUCCESSFUL
        Case Else
            MsgBox "Terjadi kesalahan pada saat merubah resolusi! Returned: " & L
    End Select
End Sub

Sub GantiWallPaper(ByVal Gbr As String)
    ret = SystemParametersInfo(SPIF_SETDESKWALLPAPER, 0&, Gbr, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Sub

Sub KosongkanWallpaper(ByVal Gambar As String)
    ret = SystemParametersInfo(SPIF_SETDESKWALLPAPER, 0&, Gambar, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Sub

Public Function ToggleScreenSaverActive(Active As Boolean) As Boolean
    Dim lActiveFlag As Long
    Dim retval As Long

    lActiveFlag = IIf(Active, 1, 0)
    retval = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, lActiveFlag, "", 0)
    ToggleScreenSaverActive = retval > 0
End Function

