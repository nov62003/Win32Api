Attribute VB_Name = "Network"
Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long


