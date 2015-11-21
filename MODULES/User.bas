Attribute VB_Name = "UserAccount"
Public vbsobj As Object
Public Get_Reg_Path, Get_Set_User_Picture_Path As String
Public Username_Exist As Boolean
Public Acc_Picture_Change As Boolean
Public Admin_Guest As Boolean
Public Username As String
Public Home As String
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function InitCommonControls Lib "COMCTL32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function PathMakeSystemFolder Lib "shlwapi.dll" Alias "PathMakeSystemFolderA" (ByVal pszPath As String) As Long
Public Const Windows_Version = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName"
Public Const Gettinghomedrive = "HKEY_CURRENT_USER\Volatile Environment\HOMEDRIVE"


Public Sub Get_Path()
Set vbsobj = CreateObject("Wscript.Shell")
    Get_Reg_Path = vbsobj.Regread("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\Common AppData")
    Get_Set_User_Picture_Path = Get_Reg_Path & "\Microsoft\User Account Pictures"
End Sub

Public Sub Username_Exist_Message()
    MsgBox "An account named '" & Username & "' already exists. Type a different name.", vbExclamation
End Sub

Public Sub Select_Control()
On Error Resume Next
    frmUserAcc.txtUsername.SetFocus
    SendKeys "{Home}+{End}"
End Sub
Public Sub Invalid_Username_Message()
    MsgBox "The specified account name is not valid, because account names cannot contain the" & vbCrLf & _
    "following characters: " & frmUserAcc.txtInvalidCha & vbCrLf & vbCrLf & _
    "Please type a different name.", vbExclamation
End Sub
Public Sub Get_Windows_Version()
On Error GoTo Err_Getting_Win_Version
Set vbsobj = CreateObject("Wscript.Shell")
    If vbsobj.Regread(Windows_Version) <> "Microsoft Windows XP" Then
        MsgBox "User Accounts Manager is designed to run only on Windows XP." & vbCrLf & "User Accounts Manager can not Continue...", vbCritical
        End
    End If
Exit Sub

Err_Getting_Win_Version:
    MsgBox "User Accounts Manager is designed to run only on Windows XP." & vbCrLf & "User Accounts Manager can not Continue...", vbCritical
    End
End Sub

Public Sub HomeDrive()
    Home = vbsobj.Regread(Gettinghomedrive)
End Sub
