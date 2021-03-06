Attribute VB_Name = "Explorer"
Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4

Const READONLY = &H1
Const HIDDEN = &H2
Const SYSTEM = &H4
Const DIRECTORY = &H10
Const ARCHIVE = &H20
Const NORMAL = &H80
Const COMPRESSED = &H800

Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long



Dim fileattrs As Long ' atribut file
Dim retval As Long ' return value

Sub aturAtribut(ByVal fa As Long, ByVal nF As String)
    Select Case fa
        Case 0
            fileattrs = FILE_ATTRIBUTE_ARCHIVE
        Case 1
            fileattrs = FILE_ATTRIBUTE_HIDDEN
        Case 2
            fileattrs = FILE_ATTRIBUTE_NORMAL
        Case 3
            fileattrs = FILE_ATTRIBUTE_READONLY
        Case 4
            fileattrs = FILE_ATTRIBUTE_SYSTEM
    End Select
    
    retval = SetFileAttributes(nF, fileattrs)
End Sub


Function DapatAtribut(ByVal NFile As String)
    Dim val As String, attr As Long
    
    attr = GetFileAttributes(NFile)
    If (attr And &H1) = &H1 Then
        val = " Read Only,"
    End If
    If (attr And &H2) = &H2 Then
        val = val & " Hidden,"
    End If
    If (attr And &H4) = &H4 Then
        val = val & " System,"
    End If
    If (attr And &H20) = &H20 Then
        val = val & " Archive,"
    End If
    If (attr And &H80) = &H80 Then
        val = val & " Normal,"
    End If
    If (attr And &H800) = &H800 Then
        val = val & " Compressed,"
    End If
    
    val = Left(val, Len(val) - 1)
    If (attr And &H10) = &H10 Then
        MsgBox "Given directory has " & val & " attributes"
    Else
        MsgBox "Given file has " & val & " attributes"
    End If

End Function

