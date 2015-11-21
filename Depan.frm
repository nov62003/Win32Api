VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDepan 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4800
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   135
      Left            =   120
      TabIndex        =   0
      Top             =   1050
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading, Silakan Tunggu..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   1400
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lingkungan Windows dengan Win32Api"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   525
      Width           =   4740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aplikasi Pengontrolan"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   225
      Width           =   2610
   End
End
Attribute VB_Name = "frmDepan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Label1.Left = (frmDepan.Width - Label1.Width) / 2
    Label2.Left = (frmDepan.Width - Label2.Width) / 2
    Label3.Left = (frmDepan.Width - Label3.Width) / 2
    PBar.Left = (frmDepan.Width - PBar.Width) / 2
End Sub

Private Sub Timer1_Timer()
    If PBar.Value < 45 Then
        PBar.Value = PBar.Value + 1
    ElseIf PBar.Value = 45 Then
        PBar.Value = 67
    ElseIf PBar.Value >= 67 And PBar.Value < 87 Then
        PBar.Value = PBar.Value + 4
        Timer1.Interval = 250
    ElseIf PBar.Value = 87 Then
        PBar.Value = 91
    ElseIf PBar.Value = 91 Then
        Timer1.Interval = 550
        PBar.Value = 97
    ElseIf PBar.Value = 97 Then
        PBar.Value = 100
    ElseIf PBar.Value = 100 Then
        Timer1.Enabled = False
        frmKontrol.Show
        Unload frmDepan
        Set frmDepan = Nothing
    End If
End Sub

