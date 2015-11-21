VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmKontrol 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "--- Aplikasi Kontrol Lingkungan Windows"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameD 
      ForeColor       =   &H00C0C0C0&
      Height          =   6015
      Left            =   2040
      TabIndex        =   13
      Top             =   -50
      Width           =   7730
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
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   2235
         Width           =   2610
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
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   4740
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0C0C0&
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   -50
      Width           =   1935
      Begin VB.CommandButton btnProses 
         Caption         =   "Keluar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5520
         Width           =   1695
      End
      Begin VB.CommandButton btnProses 
         Caption         =   "Log Off"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CommandButton btnProses 
         Caption         =   "Restart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4560
         Width           =   1695
      End
      Begin VB.CommandButton btnProses 
         Caption         =   "Shut Down"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4080
         Width           =   1695
      End
      Begin VB.CommandButton btnProses 
         Caption         =   "Back to Home"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3600
         Width           =   1695
      End
      Begin VB.CommandButton btnProses 
         Caption         =   "User Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton btnProses 
         Caption         =   "Taskbar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton btnProses 
         Caption         =   "Printer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton btnProses 
         Caption         =   "Network"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton btnProses 
         Caption         =   "Explorer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton btnProses 
         Caption         =   "Dekstop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton btnProses 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame15 
      ForeColor       =   &H00C0C0C0&
      Height          =   6015
      Left            =   2045
      TabIndex        =   61
      Top             =   -50
      Visible         =   0   'False
      Width           =   7730
      Begin VB.Frame Frame16 
         Caption         =   "Accounts Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4845
         Left            =   120
         TabIndex        =   63
         Top             =   720
         Width           =   7455
         Begin VB.Frame Frame17 
            Caption         =   "Change Passwords"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Left            =   3840
            TabIndex        =   66
            Top             =   1800
            Width           =   3495
            Begin VB.ComboBox cmbModUser 
               Height          =   315
               Left            =   1080
               Style           =   2  'Dropdown List
               TabIndex        =   88
               Top             =   240
               Width           =   2325
            End
            Begin VB.Frame Frame18 
               Caption         =   "Password"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1575
               Left            =   120
               TabIndex        =   87
               Top             =   600
               Width           =   3255
               Begin VB.TextBox txtModPW1 
                  Height          =   285
                  IMEMode         =   3  'DISABLE
                  Left            =   840
                  MaxLength       =   40
                  PasswordChar    =   "*"
                  TabIndex        =   92
                  Top             =   360
                  Width           =   2295
               End
               Begin VB.TextBox txtModPW2 
                  Height          =   285
                  IMEMode         =   3  'DISABLE
                  Left            =   840
                  MaxLength       =   40
                  PasswordChar    =   "*"
                  TabIndex        =   91
                  Top             =   840
                  Width           =   2295
               End
               Begin VB.CheckBox chkShowPW2 
                  Caption         =   "Show Password"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   90
                  Top             =   1200
                  Width           =   1815
               End
               Begin VB.Label Label12 
                  Caption         =   "New"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   94
                  Top             =   360
                  Width           =   495
               End
               Begin VB.Label Label11 
                  Caption         =   "Confirm"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   93
                  Top             =   840
                  Width           =   735
               End
            End
            Begin VB.CommandButton cmdChange 
               Caption         =   "&Change Password"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               TabIndex        =   86
               Top             =   2280
               Width           =   1695
            End
            Begin VB.Label Label13 
               Caption         =   "User name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   89
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame19 
            Caption         =   "Delete Accounts"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   3840
            TabIndex        =   65
            Top             =   240
            Width           =   3495
            Begin VB.CommandButton cmdDelete 
               Caption         =   "&Delete"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1560
               TabIndex        =   84
               Top             =   960
               Width           =   1815
            End
            Begin VB.OptionButton optKeepFiles 
               Caption         =   "&Keep Files"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   83
               Top             =   840
               Width           =   1215
            End
            Begin VB.OptionButton optDeleteFiles 
               Caption         =   "&Delete Files"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   1080
               Width           =   1335
            End
            Begin VB.ComboBox cmbDelUser 
               Height          =   315
               Left            =   1080
               TabIndex        =   81
               Top             =   360
               Width           =   2295
            End
            Begin VB.Label Label14 
               Caption         =   "User name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   85
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Create Accounts"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4455
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   3615
            Begin VB.Frame Frame21 
               Caption         =   "Type"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   165
               TabIndex        =   70
               Top             =   240
               Width           =   2055
               Begin VB.PictureBox Picture4 
                  BorderStyle     =   0  'None
                  Height          =   615
                  Left            =   40
                  ScaleHeight     =   615
                  ScaleWidth      =   1935
                  TabIndex        =   71
                  Top             =   240
                  Width           =   1935
                  Begin VB.OptionButton optLimited 
                     Caption         =   "&Limited"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Left            =   120
                     TabIndex        =   73
                     Top             =   360
                     Width           =   1215
                  End
                  Begin VB.OptionButton optAdmin 
                     Caption         =   "&Administrator"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   375
                     Left            =   120
                     TabIndex        =   72
                     Top             =   0
                     Width           =   1455
                  End
               End
            End
            Begin VB.TextBox txtUsername 
               Height          =   285
               Left            =   1080
               MaxLength       =   20
               TabIndex        =   69
               Top             =   1920
               Width           =   2415
            End
            Begin VB.Frame Frame22 
               Caption         =   "Password (optional)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1575
               Left            =   165
               TabIndex        =   68
               Top             =   2280
               Width           =   3375
               Begin VB.CheckBox chkShowPW1 
                  Caption         =   "Show Password"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   78
                  Top             =   1200
                  Width           =   2175
               End
               Begin VB.TextBox txtUserPW2 
                  Height          =   285
                  IMEMode         =   3  'DISABLE
                  Left            =   840
                  MaxLength       =   40
                  PasswordChar    =   "*"
                  TabIndex        =   77
                  Top             =   840
                  Width           =   2415
               End
               Begin VB.TextBox txtUserPW1 
                  Height          =   285
                  IMEMode         =   3  'DISABLE
                  Left            =   840
                  MaxLength       =   40
                  PasswordChar    =   "*"
                  TabIndex        =   76
                  Top             =   360
                  Width           =   2415
               End
               Begin VB.Label Label15 
                  Caption         =   "Confirm"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   80
                  Top             =   840
                  Width           =   735
               End
               Begin VB.Label Label16 
                  Caption         =   "New"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   79
                  Top             =   360
                  Width           =   495
               End
            End
            Begin VB.CommandButton cmdCreate 
               Caption         =   "&Create"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   720
               TabIndex        =   67
               Top             =   3960
               Width           =   2175
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "User name"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   75
               Top             =   1920
               Width           =   915
            End
            Begin VB.Image Image1 
               Appearance      =   0  'Flat
               Height          =   720
               Left            =   2640
               Picture         =   "Kontrol.frx":0000
               Top             =   480
               Width           =   720
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Windows accounts available."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   840
               TabIndex        =   74
               Top             =   1440
               Width           =   2535
            End
         End
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   0
         X2              =   7695
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER ACCOUNT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   120
         TabIndex        =   62
         Top             =   195
         Width           =   2430
      End
   End
   Begin VB.Frame Frame13 
      ForeColor       =   &H00C0C0C0&
      Height          =   6015
      Left            =   2045
      TabIndex        =   58
      Top             =   -50
      Visible         =   0   'False
      Width           =   7730
      Begin VB.Frame Frame23 
         Caption         =   "Ubah Nama Tombol Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   99
         Top             =   1920
         Width           =   7455
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   120
            TabIndex        =   102
            Top             =   360
            Width           =   4335
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Ganti Nama tombol Start"
            Height          =   375
            Left            =   120
            TabIndex        =   101
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Normalkan Tombol Start"
            Height          =   375
            Left            =   2400
            TabIndex        =   100
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.Frame Frame14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   59
         Top             =   720
         Width           =   7455
         Begin VB.OptionButton Option2 
            Caption         =   "Tampilkan Start"
            Height          =   375
            Left            =   120
            TabIndex        =   98
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Sembunyikan Start"
            Height          =   375
            Left            =   2040
            TabIndex        =   97
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton optSembunyi 
            Caption         =   "Sembunyikan Taksbar"
            Height          =   375
            Left            =   2040
            TabIndex        =   96
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optTampil 
            Caption         =   "Tampilkan Taksbar"
            Height          =   375
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TASKBAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   120
         TabIndex        =   60
         Top             =   195
         Width           =   1425
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   0
         X2              =   7695
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Frame Frame11 
      ForeColor       =   &H00C0C0C0&
      Height          =   6015
      Left            =   2045
      TabIndex        =   55
      Top             =   -50
      Visible         =   0   'False
      Width           =   7730
      Begin VB.Frame Frame12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   56
         Top             =   840
         Width           =   7455
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "Maaf, Bagian ini belum Selesai dikerjakan, UNDER CONSTRUCTION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   720
            Left            =   960
            TabIndex        =   103
            Top             =   240
            Width           =   5475
         End
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   0
         X2              =   7695
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRINTER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   120
         TabIndex        =   57
         Top             =   195
         Width           =   1335
      End
   End
   Begin VB.Frame Frame8 
      ForeColor       =   &H00C0C0C0&
      Height          =   6015
      Left            =   2045
      TabIndex        =   51
      Top             =   -50
      Visible         =   0   'False
      Width           =   7730
      Begin VB.Frame Frame9 
         Caption         =   "Map Network"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   52
         Top             =   840
         Width           =   7455
         Begin VB.CommandButton Command2 
            Caption         =   "Disconnect Network Drive"
            Height          =   375
            Left            =   3720
            TabIndex        =   104
            Top             =   360
            Width           =   3495
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Map Network Drive"
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NETWORK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   120
         TabIndex        =   54
         Top             =   195
         Width           =   1575
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   0
         X2              =   7695
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Frame Frame7 
      ForeColor       =   &H00C0C0C0&
      Height          =   6015
      Left            =   2045
      TabIndex        =   41
      Top             =   -50
      Visible         =   0   'False
      Width           =   7730
      Begin VB.DriveListBox drvList 
         Height          =   315
         Left            =   120
         TabIndex        =   113
         Top             =   1560
         Width           =   855
      End
      Begin VB.DirListBox dirList 
         Height          =   2115
         Left            =   120
         TabIndex        =   112
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ComboBox cboExtension 
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1080
         TabIndex        =   111
         Top             =   1560
         Width           =   855
      End
      Begin VB.PictureBox Picture2 
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   7395
         TabIndex        =   106
         Top             =   720
         Width           =   7455
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   240
            TabIndex        =   110
            Top             =   360
            Width           =   7215
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblFilesCount 
            AutoSize        =   -1  'True
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   240
            TabIndex        =   109
            Top             =   0
            Width           =   75
         End
         Begin VB.Label Label20 
            Caption         =   "Selected file: Maaf, masih ada yang salah di bagian ini!!!!!"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   240
            TabIndex        =   108
            Top             =   120
            Width           =   6375
         End
         Begin VB.Label lblShow1 
            AutoSize        =   -1  'True
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   1560
            TabIndex        =   107
            Top             =   240
            Width           =   75
         End
      End
      Begin VB.CommandButton Command3 
         Height          =   495
         Left            =   1080
         Picture         =   "Kontrol.frx":1B42
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   4200
         Width           =   855
      End
      Begin VB.Frame Frame10 
         Caption         =   "Ubah Atribut File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   42
         Top             =   4800
         Width           =   7455
         Begin VB.TextBox txtFile 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Text            =   "Maaf, masih ada yang salah di bagian ini!!!!!"
            Top             =   240
            Width           =   5775
         End
         Begin VB.CommandButton cmdFile 
            Caption         =   "Cari File"
            Height          =   375
            Left            =   6000
            TabIndex        =   43
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Archive"
            Height          =   375
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Read Only"
            Height          =   375
            Left            =   1080
            TabIndex        =   47
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Hidden"
            Height          =   375
            Left            =   2280
            TabIndex        =   48
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox Check4 
            Caption         =   "System"
            Height          =   375
            Left            =   3240
            TabIndex        =   49
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Normal"
            Height          =   375
            Left            =   4920
            TabIndex        =   50
            Top             =   600
            Width           =   975
         End
      End
      Begin MSComDlg.CommonDialog Dialog 
         Left            =   2160
         Top             =   1800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   3255
         Left            =   2040
         TabIndex        =   114
         Top             =   1560
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   5741
         View            =   1
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   0
         X2              =   7695
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EXPLORER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   120
         TabIndex        =   44
         Top             =   195
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00C0C0C0&
      Height          =   6015
      Left            =   2045
      TabIndex        =   18
      Top             =   -50
      Visible         =   0   'False
      Width           =   7730
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3240
         TabIndex        =   39
         Top             =   2160
         Width           =   4335
         Begin VB.CommandButton cmdKsng 
            Caption         =   "Kosongkan Tempat Sampah"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Desktop Icons"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   36
         Top             =   2160
         Width           =   2895
         Begin VB.OptionButton optOn 
            Caption         =   "Tampilkan"
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optOff 
            Caption         =   "Sembunyikan"
            Height          =   375
            Left            =   1440
            TabIndex        =   37
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Masukkan Password untuk mengunci Desktop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   7455
         Begin VB.CommandButton cmdBtl 
            Caption         =   "Batal"
            Height          =   375
            Left            =   6000
            TabIndex        =   35
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdBuka 
            Caption         =   "Buka Kunci"
            Height          =   375
            Left            =   4560
            TabIndex        =   34
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtPwd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   120
            PasswordChar    =   "*"
            TabIndex        =   33
            Top             =   240
            Width           =   7215
         End
         Begin VB.CommandButton cmdKunci 
            Caption         =   "Kunci"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEKSTOP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   120
         TabIndex        =   19
         Top             =   195
         Width           =   1455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   0
         X2              =   7695
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Frame Frame0 
      ForeColor       =   &H00C0C0C0&
      Height          =   6015
      Left            =   2045
      TabIndex        =   16
      Top             =   -50
      Visible         =   0   'False
      Width           =   7730
      Begin VB.Frame Frame3 
         Caption         =   "Screen Saver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   28
         Top             =   3960
         Width           =   7455
         Begin VB.OptionButton optNonAktif 
            Caption         =   "Non- Aktifkan"
            Height          =   375
            Left            =   1200
            TabIndex        =   30
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optAktif 
            Caption         =   "Aktifkan"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Background"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   3120
         TabIndex        =   23
         Top             =   840
         Width           =   4455
         Begin VB.TextBox txtCari 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   2520
            Width           =   3735
         End
         Begin MSComDlg.CommonDialog cd1 
            Left            =   3840
            Top             =   960
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton cmdKosong 
            Caption         =   "Kosongkan"
            Height          =   375
            Left            =   2640
            TabIndex        =   26
            Top             =   2040
            Width           =   1695
         End
         Begin VB.CommandButton cmdGanti 
            Caption         =   "Ganti Background"
            Height          =   375
            Left            =   2640
            TabIndex        =   25
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton cmdCari 
            Caption         =   "..."
            Height          =   375
            Left            =   3960
            TabIndex        =   24
            Top             =   2520
            Width           =   375
         End
         Begin VB.Image imgBG 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   2175
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame f 
         Caption         =   "Resolusi dan Warna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   2895
         Begin VB.CommandButton cmdAtur 
            Caption         =   "Atur Resolusi"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   2520
            Width           =   2655
         End
         Begin VB.ListBox lstResolusi 
            Appearance      =   0  'Flat
            Height          =   2175
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   0
         X2              =   7695
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DISPLAY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   120
         TabIndex        =   17
         Top             =   195
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmKontrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lNumModes As Long

Dim DelCancel As Boolean
Dim Invalid_Username As Boolean
Dim KeepFiles As Boolean
Dim Profile_Folder As Boolean
Dim FileExist As Boolean
Dim Exist_My_Doc As Boolean
Dim Exist_Desktop As Boolean
Dim CurrentAccountDeletion As Boolean
'Dim Check_Second_Folder As Boolean
Dim Err_Occurred As Boolean
'Dim Folder As String
Dim AccprofilePath
Dim DesktopFolderName, AlternateFolderName
Dim fs, DesktopPath, AlternatePath
Dim CurrentUser, NewUser As String
Dim str1, str2, str3, str4, str5, str6, str7, str8, str9, str10, str11, str12, str13, str14, str15, str16, str17, str18, str19, Str20, Str21, Str22, Str23 As String
Dim Invalid_Cha(15) As String
Const C3 = "localgroup"
Const C2 = "user"
Const C4 = "administrators"
Const C5 = "/add"
Const C1 = "net"
Const C6 = "/delete"
Const UNPATH = "Docume~1"
Const CEXE = "cmd"
Const SW_SHOW = 5
Const Administrator = "Administrator"
Const Guest = "Guest"
Const Checkforpermission = "HKEY_LOCAL_MACHINE\SOFTWARE\UAManager\Permission"
Const PrevInstance = "HKEY_LOCAL_MACHINE\SOFTWARE\UAManager\PrevInstance"
Const Logonusername = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Logon User Name"
Const Regtemppath = "HKEY_CURRENT_USER\Software\UAManager\Temp"

Private Sub chkShowPW1_Click()
    If chkShowPW1.Value = 1 Then
        txtUserPW1.PasswordChar = ""
        txtUserPW2.PasswordChar = ""
    Else
        txtUserPW1.PasswordChar = "*"
        txtUserPW2.PasswordChar = "*"
    End If
End Sub

Private Sub chkShowPW2_Click()
    If chkShowPW2.Value = 1 Then
        txtModPW1.PasswordChar = ""
        txtModPW2.PasswordChar = ""
    Else
        txtModPW1.PasswordChar = "*"
        txtModPW2.PasswordChar = "*"
    End If

End Sub

Private Sub cmbDelUser_Change()
    If cmbDelUser.Text = "" Then
        cmdDelete.Enabled = False
    Else
        cmdDelete.Enabled = True
    End If
End Sub

Private Sub cmbDelUser_Click()
cmbDelUser_Change
End Sub

Private Sub cmdChange_Click()
    If txtModPW1.Text <> "" And txtModPW2.Text <> "" Then
        If txtModPW1.Text <> txtModPW2.Text Then
            MsgBox "Password confirmation failed," & vbCrLf & "Reenter the password...!", vbExclamation
            txtModPW1.Text = ""
            txtModPW2.Text = ""
            txtModPW1.SetFocus
            Exit Sub
        End If
    End If
    If txtModPW1.Text <> "" And txtModPW2.Text = "" Then
        MsgBox "Both fields must be filled...!", vbExclamation
        txtModPW2.SetFocus
        Exit Sub
    End If
    
Change_Pword
Clear_Fields
    MsgBox "Password of '" & cmbModUser & "' changed." & vbCrLf & _
    vbCrLf & Chr(9) & "Tip:" & vbCrLf & _
    Chr(9) & "----" & vbCrLf & _
    "Incorrect User names are ignored.", vbInformation
End Sub

Private Sub cmdChangePicture_Click()
    Acc_Picture_Change = True
    frmAccPic.Show 1
    MsgBox "Account picture of '" & cmbModUser & "' changed.", vbInformation
End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdCreate_Click()
    If txtUserPW1.Text <> "" And txtUserPW2.Text <> "" Then
        If txtUserPW1.Text <> txtUserPW2.Text Then
            MsgBox "Password confirmation failed," & vbCrLf & "Reenter the password...!", vbExclamation
            txtUserPW1.Text = ""
            txtUserPW2.Text = ""
            txtUserPW1.SetFocus
            Exit Sub
        End If
    End If
    If txtUserPW1.Text <> "" And txtUserPW2.Text = "" Then
        MsgBox "Both fields must be filled...!", vbExclamation
        txtUserPW2.SetFocus
        Exit Sub
    End If

Check_Invalid_Username
    If Invalid_Username = True Then
        Invalid_Username = False
        Invalid_Username_Message
        Select_Control
        Exit Sub
    End If

    If optAdmin.Value = True And txtUserPW2 = "" Then
        Check_User_Name
            If Err_Occurred = True Then
                Err_Occurred = False
                Exit Sub
            End If

            If Username_Exist = True Then
                Username_Exist = False
                Username_Exist_Message
                Username = ""
                Select_Control
                Exit Sub
            Else
                NewUser = txtUsername
                Admin
                Clear_Fields
                Form_Activate
            End If
    End If
    If optAdmin.Value = True And txtUserPW2 <> "" Then
        Check_User_Name
            If Err_Occurred = True Then
                Err_Occurred = False
                Exit Sub
            End If

            If Username_Exist = True Then
                Username_Exist = False
                Username_Exist_Message
                Username = ""
                Select_Control
                Exit Sub
            Else
                NewUser = txtUsername
                AdminWithPW
                Clear_Fields
                Form_Activate
            End If
    End If
    If optLimited.Value = True And txtUserPW2 = "" Then
        Check_User_Name
            If Err_Occurred = True Then
                Err_Occurred = False
                Exit Sub
                End If

            If Username_Exist = True Then
                Username_Exist = False
                Username_Exist_Message
                Username = ""
                Select_Control
                Exit Sub
            Else
                NewUser = txtUsername
                Limeted
                Clear_Fields
                Form_Activate
            End If
    End If
    If optLimited.Value = True And txtUserPW2 <> "" Then
        Check_User_Name
            If Err_Occurred = True Then
                Err_Occurred = False
                Exit Sub
            End If

            If Username_Exist = True Then
                Username_Exist = False
                Username_Exist_Message
                Username = ""
                Select_Control
                Exit Sub
            Else
                NewUser = txtUsername
                LimetedWithPW
                Clear_Fields
                Form_Activate
            End If
    End If
MsgBox "Account named '" & NewUser & "' created.", vbInformation
NewUser = ""
Get_User_Names
AccountCount_Indicate
End Sub

Public Sub AdminWithPW()
On Error Resume Next
    frmAccPic.Show 1
    Shell CEXE & " " & "/c" & " " & C1 & " " & C2 & " " & txtUsername.Text & " " & txtUserPW2.Text & " " & C5, vbHide
    Shell CEXE & " " & "/c" & " " & C1 & " " & C3 & " " & C4 & " " & txtUsername.Text & " " & C5, vbHide
End Sub
Public Sub LimetedWithPW()
On Error Resume Next
    frmAccPic.Show 1
    Shell CEXE & " " & "/c" & " " & C1 & " " & C2 & " " & txtUsername.Text & " " & txtUserPW2.Text & " " & C5, vbHide
End Sub

Public Sub Admin()
On Error Resume Next
    frmAccPic.Show 1
    Shell CEXE & " " & "/c" & " " & C1 & " " & C2 & " " & txtUsername.Text & " " & C5, vbHide
    Shell CEXE & " " & "/c" & " " & C1 & " " & C3 & " " & C4 & " " & txtUsername.Text & " " & C5, vbHide
End Sub
Public Sub Limeted()
On Error Resume Next
    frmAccPic.Show 1
    Shell CEXE & " " & "/c" & " " & C1 & " " & C2 & " " & txtUsername.Text & " " & C5, vbHide
End Sub
Public Sub Delete_Account()
Dim Msg As String
Get_Current_User_Name
    If InStr(1, CurrentUser, cmbDelUser, vbTextCompare) And Len(Trim(cmbDelUser)) + 1 = Len(CurrentUser) Then
        Msg = MsgBox("It is highly recommended not to delete the Current Account." & vbCrLf & vbCrLf & _
        "DELETING THE CURRENT ACCOUNT MAY CAUSE SOME PROBLEMS." & vbCrLf & vbCrLf & _
        "Do you really need to delete the Account?", vbYesNo + vbInformation)
            If Msg = vbYes Then
                DelCancel = False
                CurrentAccountDeletion = True
                On Error Resume Next
                vbsobj.Regwrite (Regtemppath), "temp"
                Delete_Pro
            Else
                DelCancel = True
            End If
       
    Else
        Msg = MsgBox("Are you sure you need to delete this Account.", vbYesNo + vbInformation)
            If Msg = vbYes Then
                DelCancel = False
                Delete_Pro
            Else
                DelCancel = True
            End If
    End If
End Sub

Private Sub cmdDelete_Click()
Check_Administrator_Guest
    If Admin_Guest = True Then
        Admin_Guest = False
        cmbDelUser.SetFocus
        Exit Sub
    End If
Delete_Account
    'Clear_Fields
    If DelCancel = True Then
        'MsgBox "Deletion canceled...", vbInformation
        DelCancel = False
        cmbDelUser.SetFocus
    Else
        If KeepFiles = True And Profile_Folder = True Then
            If Not CurrentAccountDeletion = True Then
                KeepFiles = False
                Profile_Folder = False
                    If MsgBox("The contents of " & cmbDelUser & "'s desktop and 'My documents'" & vbCrLf & _
                    "have been saved on your Desktop in a folder called '" & cmbDelUser & "'." & vbCrLf & _
                    vbCrLf & "Do you need to open the folder?" & vbCrLf & _
                    vbCrLf & "Tip:" & vbCrLf & "----" & vbCrLf & _
                    "Incorrect User names are ignored.", vbYesNo + vbInformation) = vbYes Then Shell "Explorer " & DesktopFolderName, vbNormalFocus
                  
            ElseIf CurrentAccountDeletion = True Then
                KeepFiles = False
                Profile_Folder = False
                CurrentAccountDeletion = False
                HomeDrive
                    If MsgBox("The contents of " & cmbDelUser & "'s desktop and 'My documents'" & vbCrLf & _
                    "have been saved in drive " & Home & " in a folder called '" & cmbDelUser & "'." & vbCrLf & _
                    vbCrLf & "Do you need to open the folder?" & vbCrLf & _
                    vbCrLf & "Tip:" & vbCrLf & "----" & vbCrLf & _
                "Incorrect User names are ignored.", vbYesNo + vbInformation) = vbYes Then Shell "Explorer " & AlternateFolderName, vbNormalFocus
            End If
        Else
            MsgBox "Account '" & cmbDelUser & "' deleted." & vbCrLf & _
            vbCrLf & Chr(9) & "Tip:" & vbCrLf & _
            Chr(9) & "----" & vbCrLf & _
            "Incorrect User names are ignored.", vbInformation
        End If
    End If
Get_User_Names
cmbDelUser.SetFocus
AccountCount_Indicate
End Sub

Private Sub btnProses_Click(Index As Integer)
    Select Case Index
    Case 0  'Display
        FrameD.Visible = False
        
        Frame0.Visible = True
        Frame1.Visible = False
        Frame7.Visible = False
        Frame8.Visible = False
        Frame11.Visible = False
        Frame13.Visible = False
        Frame15.Visible = False
        
        Dim L As Long, lMaxModes As Long
        Dim lBits As Long, lWidth As Long, lHeight As Long
        
        lBits = GetDeviceCaps(hdc, BITSPIXEL)
        lWidth = Screen.Width \ Screen.TwipsPerPixelX
        lHeight = Screen.Height \ Screen.TwipsPerPixelY
        lMaxModes = 8
        
        ReDim d(0 To lMaxModes) As DEVMODE
        lNumModes = 0
        L = EnumDisplaySettings(ByVal 0, lNumModes, d(lNumModes))
        Do While L
            lstResolusi.AddItem d(lNumModes).dmPelsWidth & "x" & d(lNumModes).dmPelsHeight & "x" & d(lNumModes).dmBitsPerPel
            If lBits = d(lNumModes).dmBitsPerPel And _
                lWidth = d(lNumModes).dmPelsWidth And _
                lHeight = d(lNumModes).dmPelsHeight Then
                lstResolusi.ListIndex = lstResolusi.NewIndex
            End If
            lNumModes = lNumModes + 1
            If lNumModes > lMaxModes Then
                lMaxModes = lMaxModes + 8
                ReDim Preserve d(0 To lMaxModes) As DEVMODE
            End If
            L = EnumDisplaySettings(ByVal 0, lNumModes, d(lNumModes))
        Loop
        lNumModes = lNumModes - 1
        
        
    Case 1  'Dekstop
        FrameD.Visible = False
        
        Frame0.Visible = False
        Frame1.Visible = True
        Frame7.Visible = False
        Frame8.Visible = False
        Frame11.Visible = False
        Frame13.Visible = False
        Frame15.Visible = False
        
    Case 2  'Explorer
        FrameD.Visible = False
        
        Frame0.Visible = False
        Frame1.Visible = False
        Frame7.Visible = True
        Frame8.Visible = False
        Frame11.Visible = False
        Frame13.Visible = False
        Frame15.Visible = False
        
    Case 3  'Network
        FrameD.Visible = False
        
        Frame0.Visible = False
        Frame1.Visible = False
        Frame7.Visible = False
        Frame8.Visible = True
        Frame11.Visible = False
        Frame13.Visible = False
        Frame15.Visible = False
        
    Case 4  'Printer
        FrameD.Visible = False
        
        Frame0.Visible = False
        Frame1.Visible = False
        Frame7.Visible = False
        Frame8.Visible = False
        Frame11.Visible = True
        Frame13.Visible = False
        Frame15.Visible = False
        
    Case 5  'Taskbar
        FrameD.Visible = False
        
        Frame0.Visible = False
        Frame1.Visible = False
        Frame7.Visible = False
        Frame8.Visible = False
        Frame11.Visible = False
        Frame13.Visible = True
        Frame15.Visible = False
        
    Case 6  'User Account
        FrameD.Visible = False
        
        Frame0.Visible = False
        Frame1.Visible = False
        Frame7.Visible = False
        Frame8.Visible = False
        Frame11.Visible = False
        Frame13.Visible = False
        Frame15.Visible = True
    
    Case 7  'Other
        FrameD.Visible = True
        
        Frame0.Visible = False
        Frame1.Visible = False
        Frame7.Visible = False
        Frame8.Visible = False
        Frame11.Visible = False
        Frame13.Visible = False
        Frame15.Visible = False
    
    Case 8
        Dim Cinta As String
        Cinta = MsgBox("Anda ingin Shutdown ?", vbYesNo, "Nanya Nech...!!!")
        If Cinta = vbYes Then
            Shell "shutdown -s -f -t 0"
        Else
            MsgBox "Ngak Jadi Shutdown...!!!", vbInformation, "Informasi"
        End If
    
    Case 9
        Cinta = MsgBox("Anda ingin Restart ?", vbYesNo, "Nanya Nech...!!!")
        If Cinta = vbYes Then
            Printer.Restart
        Else
            MsgBox "Ngak Jadi Restart...!!!", vbInformation, "Informasi"
        End If
        
    Case 10
        Cinta = MsgBox("Anda ingin Log Off ?", vbYesNo, "Nanya Nech...!!!")
        If Cinta = vbYes Then
            Printer.LogOff
        Else
            MsgBox "Ngak Jadi Log Off...!!!", vbInformation, "Informasi"
        End If
        
    Case 11 'Keluar
        Cinta = MsgBox("Anda ingin Keluar ?", vbYesNo, "Nanya Nech...!!!")
        If Cinta = vbYes Then
            MsgBox "Terima Kasih!!!", vbInformation, "Informasi"
            Unload Me
            End
        Else
            MsgBox "Ngak Jadi Log Off...!!!", vbInformation, "Informasi"
        End If
        
    End Select
End Sub

Private Sub Check1_Click()
    Call Explorer.aturAtribut(0, txtFile.Text)
End Sub

Private Sub Check2_Click()
    Call Explorer.aturAtribut(3, txtFile.Text)
End Sub

Private Sub Check3_Click()
    Call Explorer.aturAtribut(1, txtFile.Text)
End Sub

Private Sub Check4_Click()
    Call Explorer.aturAtribut(4, txtFile.Text)
End Sub

Private Sub Check5_Click()
    Call Explorer.aturAtribut(2, txtFile.Text)
End Sub

Private Sub cmdAtur_Click()
    Dim nilai As Long
    nilai = lstResolusi.ListIndex
    Display.AturResolusi (nilai)
End Sub

Private Sub cmdBtl_Click()
    Unload Me
End Sub

Private Sub cmdBuka_Click()
    Desktop.bukakunci
    
    If txtPwd.Text = mypas Or txtPwd.Text = aa Or txtPwd.Text = bb Then
        Desktop.SystemParametersInfo 97, False, waste, 0
        Unload frmKunci
        'End
    Else
        MsgBox "Oops! Anda Keliru. Coba Lagiiiii"
        txtPwd.Text = ""
        txtPwd.SetFocus
    End If

End Sub

Private Sub cmdCari_Click()
    On Error GoTo ErrCut

    '------- To open a file ---------------------
    cd1.CancelError = True
    cd1.Filter = "Image Files (*.bmp)|*.bmp"
    cd1.ShowOpen
    
    Dim sfilename As String, str As String
    sfilename = cd1.FileName
    
    If sfilename <> "" Then
    
    str = Right(sfilename, 3)
    If str = "Bmp" Or str = "BMP" Or str = "bmp" Then
        imgBG.Picture = LoadPicture(sfilename)
    End If
    
    
    If FileLen(sfilename) > 6650000 Then
        MsgBox "This file is too large to open."
        Exit Sub
    End If
    txtCari.Text = sfilename
    
    Exit Sub
    End If
ErrCut:
    Close #1
    Exit Sub
End Sub

Private Sub cmdFile_Click()
    On Error GoTo ErrCut

    '------- To open a file ---------------------
    cd1.CancelError = True
    cd1.Filter = "All Files (*.*)|*.*"
    cd1.ShowOpen
    
    Dim sfilename As String, str As String
    sfilename = cd1.FileName
    
    If sfilename <> "" Then
    
    
    txtFile.Text = sfilename
    
    Exit Sub
    End If
ErrCut:
    Close #1
    Exit Sub
End Sub

Private Sub cmdGanti_Click()
    Call GantiWallPaper(txtCari.Text)
End Sub

Private Sub cmdKosong_Click()
    Display.KosongkanWallpaper ("(None)")
    txtCari.Text = ""
    imgBG.Picture = LoadPicture(None)
End Sub

Private Sub cmdKsng_Click()
    Desktop.KosongRB (Me.hwnd)
End Sub

Private Sub cmdKunci_Click()
    Desktop.kunci (txtPwd.Text)
    frmKunci.Visible = True
    Load frmKunci
    Unload Me
    
End Sub

Private Sub Command1_Click()
    Call WNetConnectionDialog(Me.hwnd, 1)
End Sub

Private Sub Command2_Click()
    Call WNetDisconnectDialog(Me.hwnd, 1)
End Sub

Private Sub Command3_Click()
    DlgLeft = 100
   DlgTop = 100
   DlgFilelabelText = "Right-Click"
   DlgFlags = OFN_EXPLORER Or OFN_ENABLEHOOK
   'InitOFN
   'ShowOpen
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To 11
        btnProses(i).BackColor = &HFFFFFF
    Next i
    Label1.Caption = UCase(Label1): Label2.Caption = UCase(Label2)
    Label1.Left = (Frame1.Width - Label1.Width) / 2
    Label2.Left = (Frame1.Width - Label2.Width) / 2
    
    On Error GoTo Permission
    Set vbsobj = CreateObject("Wscript.Shell")
    Check_Manifest_File
    PrevInstance_Handle
    Get_User_Names
    Get_Current_User_Name
    lblCurrentUserName = CurrentUser
    AccountCount_Indicate
    Check_Permission
    Check_Combo
    Open_Controls_Handler
    Exit Sub
Permission:
    'NO_Access
End Sub

Private Sub Form_Activate()
On Error GoTo Err
    optAdmin.Value = True
    txtUsername.SetFocus
    Exit Sub
Err:
    If Err.Number = 5 Then
        Exit Sub
    Else
        optAdmin.Value = False
        cmdClose.SetFocus
    End If
End Sub

Private Sub Form_Initialize()
On Error Resume Next
    Get_Windows_Version
    InitCommonControls
    Get_Path
    SetAttr Get_Set_User_Picture_Path, vbArchive + vbHidden + vbReadOnly + vbSystem
    'Set_Already_Open
End Sub

Private Sub optAktif_Click()
    Display.ToggleScreenSaverActive (True)
    optAktif.Value = True
    optNonAktif.Value = False
End Sub

Private Sub Option1_Click()
    Taskbar.SembunyiStart
End Sub

Private Sub Option2_Click()
    Taskbar.TampilkanStart
End Sub

Private Sub Option5_Click()
    Taskbar.Normal
End Sub

Private Sub Option6_Click()
    Taskbar.StartBaru (Text3.Text)
End Sub

Private Sub optNonAktif_Click()
    Display.ToggleScreenSaverActive (False)
    optAktif.Value = False
    optNonAktif.Value = True
End Sub

Private Sub optOff_Click()
    Desktop.SIcon
    optOn.Value = False
    optOff.Value = True
End Sub

Private Sub optOn_Click()
    Desktop.TIcon
    optOn.Value = True
    optOff.Value = False
End Sub

Private Sub lblContact_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lblContact.ForeColor = &HFF& Then
        lblContact.ForeColor = &HFF0000
    Else
        lblContact.ForeColor = &HFF&
    End If
End Sub

Private Sub optAdmin_Click()
On Error Resume Next
txtUsername.SetFocus
End Sub

Private Sub optLimited_Click()
txtUsername.SetFocus
End Sub
Private Sub cmbDelUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        No_Spaces_UN
        cmbDelUser.Text = ""
        SendKeys "{BS}"
    End If
End Sub

Private Sub optSembunyi_Click()
    Taskbar.SembunyiTaksbar
End Sub

Private Sub optTampil_Click()
    Taskbar.TampilkanTaksbar
End Sub

Private Sub txtModPW1_Change()
    If txtModPW1.Text <> "" Then
        chkShowPW2.Enabled = True
        txtModPW2.Enabled = True
        cmdChange.Enabled = True
        cmdChangePicture.Enabled = False
    Else
        chkShowPW2.Enabled = False
        txtModPW2.Enabled = False
        cmdChange.Enabled = False
        cmdChangePicture.Enabled = True
    End If
End Sub

Private Sub cmbModUser_Change()
    If cmbModUser.Text <> "" Then
        txtModPW1.Enabled = True
    Else
        txtModPW1.Enabled = False
        txtModPW1.Text = ""
        txtModPW2.Text = ""
    End If
End Sub

Private Sub cmbModUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        No_Spaces_UN
        cmbModUser.Text = ""
        SendKeys "{BS}"
    End If
End Sub

Private Sub txtModPW1_KeyPress(KeyAscii As Integer)

 If KeyAscii = 32 Then
        No_Spaces_Password
        SendKeys "{BS}"
 End If
End Sub


Private Sub txtModPW2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 32 Then
        No_Spaces_Password
        SendKeys "{BS}"
 End If
End Sub

Private Sub txtUsername_Change()
    If txtUsername.Text <> "" Then
        cmdCreate.Enabled = True
        txtUserPW1.Enabled = True
        txtUserPW2.Enabled = True
    Else
        cmdCreate.Enabled = False
        txtUserPW1.Enabled = False
        txtUserPW2.Enabled = False
    End If
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        No_Spaces_UN
        txtUsername.Text = ""
        SendKeys "{BS}"
    End If
End Sub

Private Sub txtUserPW1_Change()
    If txtUserPW1.Text <> "" Then
        txtUserPW2.Enabled = True
        chkShowPW1.Enabled = True
    Else
        txtUserPW2.Enabled = False
        chkShowPW1.Enabled = False
    End If
End Sub
Public Sub Change_Pword()
    Shell CEXE & " " & "/c" & " " & C1 & " " & C2 & " " & txtUsername.Text & " " & txtUserPW2.Text & " " & C5, vbHide
    Shell CEXE & " " & "/c" & " " & C1 & " " & C2 & " " & cmbModUser.Text & " " & txtModPW2.Text, vbHide
End Sub

Public Sub Clear_Fields()
    txtUsername.Text = ""
    txtUserPW1.Text = ""
    txtUserPW2.Text = ""
    'cmbDelUser.Text = ""
    'cmbModUser.Text = ""
    txtModPW1.Text = ""
    txtModPW2.Text = ""
End Sub
Public Sub No_Spaces_UN()
MsgBox "Spaces for User name not allowed...!", vbExclamation
End Sub
Public Sub No_Spaces_Password()
MsgBox "Spaces for Password not allowed...!", vbExclamation
End Sub
Sub XP_Style()
On Error Resume Next
str1 = frmStyle.txtStyle(0)
str2 = frmStyle.txtStyle(1)
str3 = frmStyle.txtStyle(2)
str4 = frmStyle.txtStyle(3)
str5 = frmStyle.txtStyle(4)
str6 = frmStyle.txtStyle(5)
str7 = frmStyle.txtStyle(6)
str8 = frmStyle.txtStyle(7)
str9 = frmStyle.txtStyle(8)
str10 = frmStyle.txtStyle(9)
str11 = frmStyle.txtStyle(10)
str12 = frmStyle.txtStyle(11)
str13 = frmStyle.txtStyle(12)
str14 = frmStyle.txtStyle(13)
str15 = frmStyle.txtStyle(14)
str16 = frmStyle.txtStyle(15)
str17 = frmStyle.txtStyle(16)
str18 = frmStyle.txtStyle(17)
str19 = frmStyle.txtStyle(18)
Str20 = frmStyle.txtStyle(19)
Str21 = frmStyle.txtStyle(20)
Str22 = frmStyle.txtStyle(21)
Str23 = frmStyle.txtStyle(22)
Open App.Path & "\" & App.EXEName & ".exe.manifest" For Output As #1
XP_Print
End Sub

Public Sub XP_Print()
Print #1, str1
Print #1, str2
Print #1, str3
Print #1, str4
Print #1, str5
Print #1, str6
Print #1, str7
Print #1, str8
Print #1, str9
Print #1, str10
Print #1, str11
Print #1, str12
Print #1, str13
Print #1, str14
Print #1, str15
Print #1, str16
Print #1, str17
Print #1, str18
Print #1, str19
Print #1, Str20
Print #1, Str21
Print #1, Str22
Print #1, Str23
Close #1
SetAttr App.Path & "\" & App.EXEName & ".exe.manifest", vbArchive + vbHidden + vbReadOnly + vbSystem
End Sub
Public Sub NO_Access()
MsgBox "User Accounts Manager is for Administrator use only.", vbCritical
optAdmin.Value = False
cmdCreate.Enabled = False
cmdDelete.Enabled = False
cmdChange.Enabled = False
optAdmin.Enabled = False
optLimited.Enabled = False
txtUsername.Enabled = False
txtUserPW1.Enabled = False
txtUserPW2.Enabled = False
chkShowPW1.Enabled = False
cmbDelUser.Enabled = False
optKeepFiles.Enabled = False
optDeleteFiles.Enabled = False
cmbModUser.Enabled = False
txtModPW1.Enabled = False
txtModPW2.Enabled = False
chkShowPW2.Enabled = False
End Sub

Public Sub Delete_Pro()
Dim C7 As String
C7 = "rmdir"
HomeDrive
'Home = vbsobj.Regread(Gettinghomedrive)

    If optKeepFiles.Value = True Then
        KeepFiles = True
            Set_My_Doc_Desktop_Contents
            Add_Folder_Picture
            Shell CEXE & " " & "/c" & " " & C1 & " " & C2 & " " & cmbDelUser.Text & " " & C6, vbHide
            Shell CEXE & " " & "/c" & " " & C7 & " " & "/s" & " " & "/q" & " " & Home & "\" & UNPATH & "\" & cmbDelUser, vbHide
            Remove_Picture_Pro
            'Shell CEXE & " " & "/c" & " " & C7 & " " & "/s" & " " & "/q" & " " & Home & "\" & UNPATH & "\" & cmbDelUser, vbHide
    ElseIf optDeleteFiles.Value = True Then
            Shell CEXE & " " & "/c" & " " & C1 & " " & C2 & " " & cmbDelUser.Text & " " & C6, vbHide
            Shell CEXE & " " & "/c" & " " & C7 & " " & "/s" & " " & "/q" & " " & Home & "\" & UNPATH & "\" & cmbDelUser, vbHide
            Remove_Picture_Pro
    End If
End Sub

Public Sub Remove_Picture_Pro()
On Error Resume Next
Dim Remove_Picture As String
    Get_Path
    Remove_Picture = Get_Set_User_Picture_Path & "\" & cmbDelUser.Text & ".bmp"
    SetAttr Remove_Picture, vbNormal
    Kill Remove_Picture
End Sub

Public Sub Check_User_Name()
On Error GoTo Err
    Dim fs, f, fc, s
    Dim str As String
    Get_Path
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(Get_Set_User_Picture_Path)
    Set fc = f.Files
        For Each f1 In fc
            s = f1.Name
            If Right(f1.Name, 3) = "bmp" Then
                str = Mid(f1.Name, 1, Len(f1.Name) - 4)
                    If InStr(1, str, txtUsername, vbTextCompare) Then
                        If Len(txtUsername) = Len(str) Then
                            Username_Exist = True
                            Username = str
                        End If
                        Exit Sub
                    End If
            End If
        Next
Exit Sub
Err:
Err_Occurred = True
MsgBox Err.Description, vbCritical
End Sub
Sub Get_User_Names()
Check_For_Administrator_Guest_Current_Acc_Pic
On Error GoTo Err
    Dim fs, f, fc, s
    Dim str, stradmin As String
    cmbDelUser.Clear
    cmbModUser.Clear
    Get_Path
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(Get_Set_User_Picture_Path)
    Set fc = f.Files
        For Each f1 In fc
            s = f1.Name
                If Right(f1.Name, 3) = "bmp" Then
                    str = Mid(f1.Name, 1, Len(f1.Name) - 4)
                       If str = "guest" Then
                           str = "Guest"
                       End If
                    cmbDelUser.AddItem str
                    cmbModUser.AddItem str
                End If
        Next
    On Error Resume Next
    cmbDelUser.Text = "Administrator"
    str = "Administrator"
    cmbModUser.Text = str
Exit Sub

Err:
    Err_Occurred = True
    MsgBox Err.Description, vbCritical
    End
End Sub
Public Sub Get_Current_User_Name()
    Dim sBuffer As String
    Dim lSize As Long

    sBuffer = Space(255)
    lSize = Len(sBuffer)
    GetUserName sBuffer, lSize
    
    CurrentUser = Left(sBuffer, lSize)
End Sub

Public Sub Check_Administrator_Guest()
    If InStr(1, Administrator, cmbDelUser, vbTextCompare) And Len(Administrator) = Len(cmbDelUser) Then
        MsgBox "The account 'Administrator' cannot be deleted.", vbCritical
        Admin_Guest = True
    ElseIf InStr(1, Guest, cmbDelUser, vbTextCompare) And Len(Guest) = Len(cmbDelUser) Then
        MsgBox "The account 'Guest' cannot be deleted.", vbCritical
        Admin_Guest = True
    End If
End Sub

Public Sub Check_Invalid_Username()
Invalid_Cha(0) = Mid(txtInvalidCha, 1, 1)
Invalid_Cha(1) = Mid(txtInvalidCha, 2, 1)
Invalid_Cha(2) = Mid(txtInvalidCha, 3, 1)
Invalid_Cha(3) = Mid(txtInvalidCha, 4, 1)
Invalid_Cha(4) = Mid(txtInvalidCha, 5, 1)
Invalid_Cha(5) = Mid(txtInvalidCha, 6, 1)
Invalid_Cha(6) = Mid(txtInvalidCha, 7, 1)
Invalid_Cha(7) = Mid(txtInvalidCha, 8, 1)
Invalid_Cha(8) = Mid(txtInvalidCha, 9, 1)
Invalid_Cha(9) = Mid(txtInvalidCha, 10, 1)
Invalid_Cha(10) = Mid(txtInvalidCha, 11, 1)
Invalid_Cha(11) = Mid(txtInvalidCha, 12, 1)
Invalid_Cha(12) = Mid(txtInvalidCha, 13, 1)
Invalid_Cha(13) = Mid(txtInvalidCha, 14, 1)
Invalid_Cha(14) = Mid(txtInvalidCha, 15, 1)
Invalid_Cha(15) = Mid(txtInvalidCha, 16, 1)
    For i = LBound(Invalid_Cha) To UBound(Invalid_Cha)
        If InStr(1, txtUsername, Invalid_Cha(i), vbTextCompare) Then
            Invalid_Username = True
        End If
    Next
End Sub

Public Sub AccountCount_Indicate()
lblAccCount.Caption = cmbModUser.ListCount
End Sub

Public Sub Check_Profile_Folder()
Dim fs
On Error GoTo Err
    Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.FolderExists(AccprofilePath & "\" & cmbDelUser) Then
            Profile_Folder = True
            Exit Sub
        Else
            Exit Sub
        End If
Exit Sub

Err:
    Exit Sub
End Sub
Public Sub Check_Manifest_File()
'Dim Gettingfile, fs
Dim fs
On Error GoTo Err
    Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.FileExists(App.Path & "\" & App.EXEName & ".exe.manifest") Then
            FileExist = True
            SetAttr App.Path & "\" & App.EXEName & ".exe.manifest", vbArchive + vbHidden + vbReadOnly + vbSystem
            Exit Sub
        Else
            vbsobj.Regwrite (PrevInstance), "False"
            XP_Style
            Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
            End
        End If
Exit Sub

Err:
    XP_Style
    Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
    End
    'Exit Sub
End Sub
Public Sub Open_Controls_Handler()
    cmdCreate.Enabled = False
    cmdChange.Enabled = False
    txtUserPW1.Enabled = False
    txtUserPW2.Enabled = False
    'txtModPW1.Enabled = False
    'txtModPW2.Enabled = False
    chkShowPW1.Enabled = False
    chkShowPW2.Enabled = False
    optKeepFiles.Value = True
End Sub

Public Sub Check_Combo()
    If cmbDelUser = "" Then
        cmdDelete.Enabled = False
    Else
        cmdDelete.Enabled = True
    End If
    
    If cmbModUser = "" Then
        cmdChangePicture.Enabled = False
    Else
        cmdChangePicture.Enabled = True
    End If
End Sub

Public Sub Set_My_Doc_Desktop_Contents()
    Dim Folder As String
    'Set vbsobj = CreateObject("WScript.Shell")
    Set fs = CreateObject("Scripting.FileSystemObject")
    Get_Path
    AccprofilePath = Mid(Get_Reg_Path, 1, 25)
    Check_Profile_Folder
        If Profile_Folder = True Then
            On Error Resume Next
            If CurrentAccountDeletion = True Then
                HomeDrive
                AlternatePath = Home
                AlternateFolderName = AlternatePath & "\" & cmbDelUser
                MkDir AlternateFolderName
            Else
                DesktopPath = vbsobj.SpecialFolders("Desktop")
                DesktopFolderName = DesktopPath & "\" & cmbDelUser
                MkDir DesktopFolderName
            End If
            
    Check_My_Doc_Desktop
            
        If Exist_My_Doc = True Then
            Exist_My_Doc = False
            Folder = "Copy of My Documents"
        Else
            Folder = "My Documents"
        End If
        
        If CurrentAccountDeletion = True Then
            MoveActionMyDoc = fs.MoveFolder(AccprofilePath & "\" & cmbDelUser & "\My Documents", AlternateFolderName & "\" & Folder)
        Else
            MoveActionMyDoc = fs.MoveFolder(AccprofilePath & "\" & cmbDelUser & "\My Documents", DesktopFolderName & "\" & Folder)
        End If
        
        If Exist_Desktop = True Then
            Exist_Desktop = False
            Folder = "Copy of Desktop"
        Else
            Folder = "Desktop"
        End If
        
        If CurrentAccountDeletion = True Then
            MoveActionDesktop = fs.MoveFolder(AccprofilePath & "\" & cmbDelUser & "\Desktop", AlternateFolderName & "\" & Folder)
        Else
            MoveActionDesktop = fs.MoveFolder(AccprofilePath & "\" & cmbDelUser & "\Desktop", DesktopFolderName & "\" & Folder)
        End If
        Else
        Exit Sub
        End If
End Sub

Public Sub Add_Folder_Picture()
On Error Resume Next
    If CurrentAccountDeletion = True Then
        Make_as_systemfolder = PathMakeSystemFolder(AlternateFolderName)
    Else
        Make_as_systemfolder = PathMakeSystemFolder(DesktopFolderName)
    End If
    
    Set_Folder_Icon_Info
End Sub

Public Sub Check_My_Doc_Desktop()
On Error GoTo Err
    Set fs = CreateObject("Scripting.FileSystemObject")
    If CurrentAccountDeletion = True Then
        If fs.FolderExists(AlternateFolderName & "\My Documents") Then
            Exist_My_Doc = True
        End If
        
        If fs.FolderExists(AlternateFolderName & "\Desktop") Then
            Exist_Desktop = True
        End If
        
    Else
        If fs.FolderExists(DesktopFolderName & "\My Documents") Then
            Exist_My_Doc = True
            'Folder = "My Documents - 1"
        End If
        If fs.FolderExists(DesktopFolderName & "\Desktop") Then
            Exist_Desktop = True
            'Folder = "Desktop - 1"
        End If
    End If
Exit Sub

Err:
End Sub

Public Sub Check_Permission()
    vbsobj.Regwrite (Checkforpermission), "Allow"
End Sub
Public Sub PrevInstance_Handle()
On Error Resume Next
    If vbsobj.Regread(PrevInstance) = "True" Then
        If App.PrevInstance Then
       'vbsobj.Regdelete (test)
        End
        End If
    ElseIf vbsobj.Regread(PrevInstance) = "False" Then
        vbsobj.Regwrite (PrevInstance), "True"
        Exit Sub
    End If
End Sub

Public Sub Set_Folder_Icon_Info()
On Error Resume Next
Dim l1, l2, l3
l1 = "[.ShellClassInfo]"
l2 = "IconFile=%SystemRoot%\system32\SHELL32.dll"
l3 = "IconIndex=170"
    If CurrentAccountDeletion = True Then
        Open AlternateFolderName & "\Desktop.ini" For Output As #1
    Else
        Open DesktopFolderName & "\Desktop.ini" For Output As #1
    End If

        Print #1, l1
        Print #1, l2
        Print #1, l3
        Close #1
        
    If CurrentAccountDeletion = True Then
        SetAttr AlternateFolderName & "\Desktop.ini", vbHidden + vbSystem
    Else
        SetAttr DesktopFolderName & "\Desktop.ini", vbHidden + vbSystem
    End If
End Sub

Private Sub txtUserPW1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 32 Then
        No_Spaces_Password
        SendKeys "{BS}"
 End If
End Sub

Private Sub txtUserPW2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 32 Then
        No_Spaces_Password
        SendKeys "{BS}"
 End If
End Sub

Public Sub Check_For_Administrator_Guest_Current_Acc_Pic()
Get_Path
Dim Administrator_Pic, Guest_Pic, Current_Acc_Pic As String
Dim fs
On Error GoTo Err0
    Set fs = CreateObject("Scripting.FileSystemObject")
        Administrator_Pic = Get_Set_User_Picture_Path & "\" & "Administrator.bmp"
            If Not fs.FileExists(Administrator_Pic) Then
                SavePicture frmAccPic.imgAP(1), Administrator_Pic
            End If
        
        Guest_Pic = Get_Set_User_Picture_Path & "\" & "Guest.bmp"
            If Not fs.FileExists(Guest_Pic) Then
                SavePicture frmAccPic.imgAP(2), Guest_Pic
            End If
            
        Current_Acc_Pic = Get_Set_User_Picture_Path & "\" & vbsobj.Regread(Logonusername) & ".bmp"
            On Error Resume Next
            If Not fs.FileExists(Current_Acc_Pic) Then
                If vbsobj.Regread(Regtemppath) <> "temp" Then
                    SavePicture frmAccPic.imgAP(3), Current_Acc_Pic
                End If
            End If
            
Exit Sub
    
Err0:
If Err.Number = 75 Then
MsgBox "Unexpected error occured.", vbCritical
Else
MsgBox Err.Description, vbCritical
End If
Exit Sub

End Sub


Private Function Crypt(fName As String, PW As String) As Boolean

  On Error GoTo CErr
  Dim FTemp As String, lRet As Long
  Crypt = False
  FTemp = fName & "tesseract"
  'lRet = cFileCrypt(fName, FTemp, PW)
  If lRet < 0 Then Exit Function
  Kill fName
  Name FTemp As fName
  Crypt = True
  Exit Function
CErr:
  Crypt = False
  MsgBox "Error: " & Error(Err), vbCritical, "Error"
  Exit Function
End Function

Private Sub SmartVisible(C As Control, f As Boolean)
  If (C.Visible <> f) Then C.Visible = f
End Sub

Private Sub Talk(Optional Msg)
  If fHog Then Exit Sub
  If IsMissing(Msg) Then
    If fStatB.SimpleText <> "" Then fStatB.SimpleText = ""
  Else
   ' If fStatB.SimpleText <> Msg Then fStatB.SimpleText = Msg
  End If
End Sub


Private Sub TalkHog(Optional Msg)
  If IsMissing(Msg) Then
    fStatB.SimpleText = ""
    fHog = False
  Else
    If fStatB.SimpleText <> Msg Then fStatB.SimpleText = Msg
    fHog = True
  End If
End Sub

Private Sub cboExtension_Click()
Clipboard.Clear
If Trim(cboExtension.Text) = "" Then
        MsgBox "No search pattern yet"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    Dim mFirstPath As String
    Dim mErrDirDiver As Boolean
    Dim mDirCount As Integer
    Dim mNumFiles As Integer
    
    If dirList.Path <> dirList.List(dirList.ListIndex) Then
         dirList.Path = dirList.List(dirList.ListIndex)
        
         Screen.MousePointer = vbDefault
        
    End If

    filList.Pattern = cboExtension.Text

    mFirstPath = dirList.Path
    mDirCount = dirList.ListCount

    filesCount = 0
         Screen.MousePointer = vbDefault
    If mErrDirDiver = True Then
         
         filesCount = 0
         dirList.Path = CurDir
         drvList.Drive = dirList.Path
         Screen.MousePointer = vbDefault
         Exit Sub
    End If
    If filesCount > 0 Then
    End If
    filList.Path = dirList.Path
    DirList_Change
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub DirList_Change()
       filList.Path = dirList.Path
       lblFilesCount.Caption = "Files in directory " & dirList.Path & ": " & ListView1.ListItems.Count
       lblShow1.Caption = ""

End Sub

Private Sub DirList_LostFocus()
    dirList.Path = dirList.List(dirList.ListIndex)
End Sub

Private Sub dirList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Talk "Browse folders"
End Sub


Private Sub drvList_Change()
   On Error GoTo DriveHandler
    dirList.Path = drvList.Drive
    Exit Sub

DriveHandler:
    drvList.Drive = dirList.Path
    Exit Sub
End Sub

Private Sub DirList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> vbRightButton Then
         Exit Sub
    End If
    'PopupMenu Me.popFeaturesDir
End Sub

Private Sub filList_PathChange()
  On Error Resume Next
  Dim bRet As Boolean
  DoEvents
  ListView1.ListItems.Clear
  Set ListView1.Icons = Nothing
  Set ListView1.SmallIcons = Nothing
  imgFiles.ListImages.Clear
  MousePointer = vbHourglass
  TalkHog "Retrieving files in folder: " & dirList.Path
  Dim sPath As String, w As Long
  sPath = IIf(Right(filList.Path, 1) = "\", filList.Path, filList.Path & "\")
  Dim imgT As ListImage, i As Integer, hIcon
  For i = 0 To filList.ListCount - 1
    w = 1
    'hIcon = ExtractAssociatedIcon(0, sPath & filList.List(i), w)
    If IsNull(hIcon) Then
      picBuff.Picture = picDefault.Picture
    Else
      Set picBuff.Picture = Nothing
      DoEvents
      'DrawIcon picBuff.hdc, 0, 0, hIcon
      DoEvents
      picBuff.Picture = picBuff.Image
      DoEvents
    End If
    Set imgT = imgFiles.ListImages.Add(, , picBuff.Picture)
  Next
  ListView1.Icons = imgFiles
  ListView1.SmallIcons = imgFiles
  For i = 0 To filList.ListCount - 1
    ListView1.ListItems.Add , , filList.List(i), i + 1, i + 1
  Next
  TalkHog
  MousePointer = vbDefault
End Sub

Private Sub ListView1_Click()
On Error Resume Next
lblShow1.Caption = dirList.Path & "\" & ListView1.SelectedItem.Text

End Sub

Private Sub ListView1_DblClick()
Call ShellExecute(hwnd, "Open", lblShow.Caption, "", App.Path, 1)
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim L As ListItem
  Set L = ListView1.HitTest(x, y)
  If L Is Nothing Then Exit Sub
  lblShow.Caption = dirList.Path & "\" & L.Text
  TalkHog ""
End Sub


Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button <> vbRightButton Then
         Exit Sub
    End If
    'PopupMenu Me.popFeatures
End Sub


Private Sub mnuOrderAZ_Click()
 ListView1.SortOrder = 0
    ListView1.Sorted = True
    
    mnuOrderAZ.Checked = ListView1.SortOrder = 0
    mnuOrderZA.Checked = mnuOrderAZ.Checked = False
End Sub

Private Sub mnuOrderZA_Click()
ListView1.SortOrder = 1
    ListView1.Sorted = True
    
    mnuOrderAZ.Checked = ListView1.SortOrder = 0
    mnuOrderZA.Checked = mnuOrderAZ.Checked = False
End Sub

Private Sub picSize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  DrawMode = vbXorPen
  DrawWidth = 3
  Line (picSize.Left + x, picSize.Top)-(picSize.Left + x, picSize.Top + picSize.Height), vbWhite
  fLastX = picSize.Left + x
End Sub


Private Sub picSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If fLastX >= 0 Then
    Line (fLastX, picSize.Top)-(fLastX, picSize.Top + picSize.Height), vbWhite
  Else
    Exit Sub
  End If
  Line (picSize.Left + x, picSize.Top)-(picSize.Left + x, picSize.Top + picSize.Height), vbWhite
  fLastX = picSize.Left + x
End Sub

Private Sub picSize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If fLastX >= 0 Then
    Line (fLastX, picSize.Top)-(fLastX, picSize.Top + picSize.Height), vbWhite
  End If
  fLastX = -100
  
End Sub

Private Sub popFeaturesCopy_Click()
Dim hWndDesk As Integer
   
   Dim params As String
   Dim R As Long
   Dim mPath As String, mfilespec As String
   If ListView1.SelectedItem.Selected Then
  
    mfilespec = lblShow1.Caption
      params = vbNullString
      
    Dim oldDir As String
    mPath = CurDir
    
    Dialog.FileName = mfilespec
    Dialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    Dialog.FilterIndex = 4
    Dialog.flags = cdlOFNOverwritePrompt
    Dialog.CancelError = False
    Dialog.ShowSave
    
afterCreatingDir:
    If mfilespec <> Dialog.FileName Then
        FileCopy mfilespec, Dialog.FileName
        If PathSection(mfilespec, 1) = PathSection(Dialog.FileName, 1) Then
             ListView1.Refresh
            
         End If
         ChDir mPath
    End If
    Exit Sub
    
errHandler:
    If Err.Number <> 32755 Then
         If Err <> 76 Then
             ErrMsgProc "popFeaturesCopy"
         Else
              If MsgBox("Dir " & PathSection(mfilespec, 1) & " does not exist" _
                   & vbCrLf & "Create it?", vbYesNo + vbQuestion) = vbNo Then
                   Exit Sub
              End If
              MkDir PathSection(mfilespec, 1)
              ListView1.Refresh
              GoTo afterCreatingDir
         End If
     End If
      End If
End Sub

Private Sub popFeaturesCopy2_Click()
On Error GoTo errHandler
Dim mPath As String, mfilespec As String
Dim mFullSpec As String

    mfilespec = lblShow1.Caption

    Dim oldDir As String
    mPath = CurDir
    
  
    mfilespec = lblShow1.Caption
    
    Dialog.FileName = mfilespec
    Dialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    Dialog.FilterIndex = 1
    Dialog.CancelError = False
    
again:
    Dialog.ShowSave
    
    mFullSpec = dirList.List(ListView1.SelectedItem.Selected) & Dialog.FileName
    
    If PathSection(mFullSpec, 1) = PathSection(mfilespec, 1) Then
        MsgBox "Cannot move to the same directory"
        GoTo again
    End If
    
afterCreatingDir:
    FileCopy mfilespec, Dialog.FileName
    Kill mfilespec
    DoEvents
    
    ChDir mPath
    ListView1.Refresh
    
    Exit Sub
    
errHandler:
    If Err.Number <> 32755 Then
         If Err <> 76 Then
              ErrMsgProc "popFeaturesCopy2_Click"
         Else
              If MsgBox("Directory " & vbCrLf & PathSection(Dialog.FileName, 1) & _
                   " does not exist" & vbCrLf & "Create it?", vbYesNo + _
                   vbQuestion) = vbNo Then
                   Exit Sub
              End If
              MkDir PathSection(Dialog.FileName, 1)
              dirList.Refresh
              GoTo afterCreatingDir
         End If
    End If
End Sub

Private Sub popFeaturesDelete_Click()
Dim hWndDesk As Integer
 
   Dim params As String
   Dim R As Long
   
   If ListView1.SelectedItem.Selected Then
   Dim mfilespec As String
    mfilespec = lblShow1.Caption
    
      params = vbNullString
     
    If MsgBox("Sure to delete " & lblShow.Caption & vbCrLf, _
           vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
    End If
 
    On Error GoTo errHandler
    Kill mfilespec
    DoEvents
   ListView1.SelectedItem.Selected = ""
   ListView1.SelectedItem.Selected = False
   ListView1.Refresh
  
  
      End If
      Exit Sub
errHandler:
End Sub

Private Sub popFeaturesDirCreate_Click()
On Error GoTo errHandler
    Dim currdir As String, newDir As String
    currdir = dirList.List(dirList.ListIndex)
again:
    newDir = InputBox("Type full directory specification:", _
        "Create directory", currdir)
    If newDir = "" Then
         Exit Sub
    End If
    MkDir newDir
    DoEvents
    dirList.Refresh
    Exit Sub
errHandler:
    If Err.Number = 75 Then
        MsgBox "Directory already exists/access error"
        GoTo again
    End If
    ErrMsgProc "popFeaturesDirCreate_click"
End Sub

Private Sub popFeaturesDirDelete_Click()
On Error GoTo errHandler
    If MsgBox("Sure to delete " & dirList.List(dirList.ListIndex) & vbLf & _
           "and all its contents?", vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
    End If
    Dim currdir As String, delDir As String
    currdir = CurDir
    delDir = dirList.List(dirList.ListIndex)
    ChDir delDir
    On Error Resume Next
    Kill "*.*"
    On Error GoTo errHandler
    ChDir currdir
    RmDir dirList.List(dirList.ListIndex)
    dirList.Refresh
    Exit Sub
errHandler:
    ErrMsgProc "popFeaturesDirDelete_click"
End Sub

Private Sub popFeaturesDirRename_Click()
On Error GoTo errHandler
    Dim newDir As String, origDirAsFile As String
    Dim origFullPath As String, origPathDir As String
    origFullPath = dirList.List(dirList.ListIndex)
    origPathDir = PathSection(origFullPath, 1)
    origDirAsFile = PathSection(origFullPath, 2)
    
again:
    newDir = InputBox("Type new name", "Rename directory", origDirAsFile)
    If newDir = "" Then
         Exit Sub
    End If
    If InStr(newDir, "\") <> 0 Then
         MsgBox "Rename directory within same path only"
         GoTo again
    ElseIf PathSection(origPathDir & newDir, 1) <> PathSection(origFullPath, 1) Then
          
         MsgBox "Rename file within same directory only"
         GoTo again
    End If
    
    newDir = origPathDir & newDir
    Name origFullPath As newDir
    dirList.Path = newDir
    Exit Sub

errHandler:
    ErrMsgProc "popDirListRename_click"
End Sub

Private Sub popFeaturesOpen_Click()
Dialog.DialogTitle = "Right-Click on a FILE or FOLDER to make your changes"
Dim mfilespec
mfilespec = lblShow1.Caption
Dialog.FileName = mfilespec
Dialog.Filter = "All Files (*.*)|*.*"
Dialog.ShowOpen
End Sub

Private Sub popFeaturesProperties_Click()
Dim R
'r = ShowFileProperties(lblShow1.Caption, Me.hwnd)
    If R <= 32 Then MsgBox "Error"
  
    MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub popFeaturesRename_Click()
On Error GoTo errHandler
   
    Dim newName As String, origName As String
    Dim newFileSpec As String, oldDir As String
    Dim mfilespec As String
    mfilespec = lblShow.Caption
    oldDir = PathSection(lblShow1.Caption, 1)
    origName = PathSection(lblShow1.Caption, 2)
    
again:
    newName = InputBox("Type new name including extension", "Rename file", origName)
    If newName = "" Then
         Exit Sub
    End If
    
    If InStr(newName, "\") <> 0 Then
         MsgBox "Rename file within same directory only"
         GoTo again
    ElseIf PathSection(oldDir & newName, 1) <> PathSection(lblShow.Caption, 1) Then
         
         MsgBox "Rename file within same directory only"
         GoTo again
    End If
    
    newFileSpec = oldDir & newName
    Name mfilespec As newFileSpec
    DoEvents
    ListView1.Refresh
    If mfilespec = FileListFileSpec Then
         FileListFileSpec = newFileSpec
      
    End If
    Exit Sub
    
errHandler:
  
    
    ErrMsgProc "popFeatureRename_click"
End Sub

Function PathSection(ByVal inPath As String, inReturnType As Integer)
    Dim DriveLetter As String
    Dim DirPath As String
    Dim fName As String
    Dim Extension As String
    Dim PathLength As Integer
    Dim ThisLength As Integer
    Dim Offset As Integer
    Dim FileNameFound As Boolean

    If inReturnType <> 0 And inReturnType <> 1 And inReturnType <> 2 And inReturnType <> 3 Then
        Err.Raise 1
        Exit Function
    End If
    DriveLetter = ""
    DirPath = ""
    fName = ""
    Extension = ""
    If Mid(inPath, 2, 1) = ":" Then
        DriveLetter = Left(inPath, 2)
        inPath = Mid(inPath, 3)
    End If
    PathLength = Len(inPath)
    For Offset = PathLength To 1 Step -1
        Select Case Mid(inPath, Offset, 1)
            Case ".":
            ThisLength = Len(inPath) - Offset
            If ThisLength >= 1 Then
                Extension = Mid(inPath, Offset, ThisLength + 1)
            End If
            inPath = Left(inPath, Offset - 1)
            Case "\":
            ThisLength = Len(inPath) - Offset
            If ThisLength >= 1 Then
                fName = Mid(inPath, Offset + 1, ThisLength)
                inPath = Left(inPath, Offset)
                FileNameFound = True
                Exit For
            End If
            Case Else
        End Select
    Next Offset
    If FileNameFound = False Then
        fName = inPath
    Else
        DirPath = inPath
    End If
    If inReturnType = 0 Then
        PathSection = DriveLetter
    ElseIf inReturnType = 1 Then
        PathSection = DirPath
    ElseIf inReturnType = 2 Then
        PathSection = fName & Extension
    ElseIf inReturnType = 3 Then
        PathSection = fName
    ElseIf inReturnType = 4 Then
        PathSection = Extension
    End If
End Function

Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub

Private Sub popFreaturesDirProperties_Click()
Dim fName As String
Dim R As Long
    
    fName = dirList.Path
    Caption = "[" & fName & "]"
    
    MousePointer = vbHourglass
    DoEvents
  ' r = ShowFileProperties(fName, Me.hwnd)
    If R <= 32 Then MsgBox "Error"
  
    MousePointer = vbDefault
    Exit Sub
End Sub


