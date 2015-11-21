VERSION 5.00
Begin VB.Form frmKunci 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   10950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15090
   Icon            =   "Locked.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   15090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   5700
      Picture         =   "Locked.frx":030A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   4740
      TabIndex        =   5
      Top             =   8760
      Width           =   5415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Masukkan Password untuk membuka kunci"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4980
      TabIndex        =   4
      Top             =   7080
      Width           =   4935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Created by YT. Paulus nov62003@yahoo.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   5100
      TabIndex        =   3
      Top             =   9360
      Width           =   4695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "Tekan Tombol Control dan Backspace untuk memasukkan Password"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4680
      TabIndex        =   2
      Top             =   7440
      Width           =   5535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   6240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Desktop sedang di kunci"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   5880
      TabIndex        =   0
      Top             =   5160
      Width           =   3135
   End
End
Attribute VB_Name = "frmKunci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Kk, a, R, t, h, i, K, e As Long
Private Sub Form_KeyPress(KeyAscii As Integer)
Dim wndir As String, mypath As String
Dim x As Long
If KeyAscii = 107 Then
Kk = 1
End If
If KeyAscii = 97 Then
a = 1
End If
If KeyAscii = 114 Then
R = 1
End If
If KeyAscii = 116 Then
t = 1
End If
If KeyAscii = 104 Then
h = 1
End If
If KeyAscii = 105 Then
i = 1
End If
If KeyAscii = 75 Then
K = 1
End If
If KeyAscii = 56 Then
e = 1
End If
If KeyAscii = 127 Then
frmKontrol.Visible = True
Load frmKontrol

End If
If Kk = 1 And K = 1 And a = 1 And R = 1 And t = 1 And h = 1 And i = 1 And e = 1 Then
'Enable Ctrl+Alt+Del,Ctrl+Esc and Alt+Tab
Desktop.SystemParametersInfo 97, False, waste, 0
wndir = String(128, " ")
mypath = String(128, " ")
x = Desktop.GetWindowsDirectory(wndir, 128)
wndir = Left(wndir, InStr(wndir, Chr(0)) - 1)
mypath = wndir & "\kpass.ini"
Kill (mypath)
End
End If
End Sub

Private Sub Form_Load()
Dim meuser As String, mecomp As String
meuser = String(128, " ")
mecomp = String(128, " ")
xx = Desktop.GetUserName(meuser, 128)
xy = Desktop.GetComputerName(mecomp, 128)
mecomp = Left(mecomp, InStr(mecomp, Chr(0)) - 1)
meuser = Left(meuser, InStr(meuser, Chr(0)) - 1)
'Disable Ctrl+Alt+Del,Ctrl+Esc and Alt+Tab
Desktop.SystemParametersInfo 97, True, waste, 0

'Label2.Caption = "at " & Time & ", " & Date
'Label6.Caption = "User " & meuser & " is logged in this computer named " & mecomp
End Sub

