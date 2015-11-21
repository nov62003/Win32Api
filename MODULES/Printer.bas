Attribute VB_Name = "Printer"
Dim Hasil

Const EWX_LogOff As Long = 0
Const EWX_SHUTDOWN As Long = 1
Const EWX_REBOOT As Long = 2

Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long


Sub ShutDownWindows(ByVal uFlags As Long)
   Call ExitWindowsEx(uFlags, 0)
End Sub

Sub LogOff()
    Hasil = ExitWindowsEx(EWX_LogOff, 0&)
End Sub

Sub TurnOff()
    Hasil = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Sub

Sub Restart()
    Hasil = ExitWindowsEx(EWX_REBOOT, 0&)
End Sub
