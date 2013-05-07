VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Control Panel , Swift Airlines Pvt LTD, © [ 2011 - 2015 ]"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   14010
   Icon            =   "user_panel.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "user_panel.frx":3D1A
   ScaleHeight     =   7545
   ScaleWidth      =   14010
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "booking"
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   3480
         Left            =   0
         Negotiate       =   -1  'True
         Picture         =   "user_panel.frx":280DE
         ScaleHeight     =   14.25
         ScaleMode       =   4  'Character
         ScaleWidth      =   120
         TabIndex        =   109
         Top             =   0
         Width           =   14460
         Begin VB.Label Label76 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7200
            TabIndex        =   127
            ToolTipText     =   "Auckland"
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label75 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6360
            TabIndex        =   126
            ToolTipText     =   "Singapore"
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label74 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10320
            TabIndex        =   125
            ToolTipText     =   "Capetown"
            Top             =   2760
            Width           =   255
         End
         Begin VB.Label Label73 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1320
            TabIndex        =   124
            ToolTipText     =   "New York City"
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label Label72 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6600
            TabIndex        =   123
            ToolTipText     =   "Tokyo"
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label71 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8280
            TabIndex        =   122
            ToolTipText     =   "Beijing"
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label70 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8040
            TabIndex        =   121
            ToolTipText     =   "Hong Kong"
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label Label69 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4680
            TabIndex        =   120
            ToolTipText     =   "Dubai"
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label68 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3960
            TabIndex        =   119
            ToolTipText     =   "Berlin"
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label67 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3480
            TabIndex        =   118
            ToolTipText     =   "London"
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label66 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   117
            ToolTipText     =   "Austin"
            Top             =   2280
            Width           =   255
         End
         Begin VB.Label Label65 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   480
            TabIndex        =   116
            ToolTipText     =   "San fransisco"
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label Label64 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7080
            TabIndex        =   115
            ToolTipText     =   "Sydney"
            Top             =   2640
            Width           =   255
         End
         Begin VB.Label Label63 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   9120
            TabIndex        =   114
            ToolTipText     =   "Bangalore"
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8880
            TabIndex        =   113
            ToolTipText     =   "Patna"
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5280
            TabIndex        =   112
            ToolTipText     =   "Kolkata"
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5040
            TabIndex        =   111
            ToolTipText     =   "Delhi"
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Cooper Std Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5040
            TabIndex        =   110
            ToolTipText     =   "Mumbai"
            Top             =   1560
            Width           =   255
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "GO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   86
         Top             =   6840
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2880
         TabIndex        =   83
         Top             =   5040
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   855
         Left            =   11880
         Picture         =   "user_panel.frx":41068
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6480
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8520
         TabIndex        =   13
         Top             =   6480
         Width           =   3135
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8520
         TabIndex        =   12
         Top             =   6120
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11880
         TabIndex        =   11
         Text            =   "Select Class"
         Top             =   6000
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000005&
         Caption         =   "OTHER"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10440
         TabIndex        =   10
         Top             =   7080
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000005&
         Caption         =   "FEMALE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9480
         TabIndex        =   9
         Top             =   7080
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "MALE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8640
         TabIndex        =   8
         Top             =   7080
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8520
         TabIndex        =   7
         Top             =   5760
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8520
         TabIndex        =   6
         Top             =   4920
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8520
         TabIndex        =   5
         Top             =   4440
         Width           =   3135
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2160
         Left            =   3480
         TabIndex        =   4
         Top             =   4080
         Width           =   2295
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   360
         TabIndex        =   3
         Top             =   5640
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   4920
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   4080
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3480
         TabIndex        =   84
         Top             =   6840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   107282435
         CurrentDate     =   41382
      End
      Begin VB.Label Label82 
         BackStyle       =   0  'Transparent
         Caption         =   "Fare : Rs 100000"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11880
         TabIndex        =   108
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label81 
         BackStyle       =   0  'Transparent
         Caption         =   "(Business Class )"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11880
         TabIndex        =   107
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label Label80 
         BackStyle       =   0  'Transparent
         Caption         =   "Fare : Rs 100000"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11880
         TabIndex        =   106
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Label79 
         BackStyle       =   0  'Transparent
         Caption         =   "( Economy Class )"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11880
         TabIndex        =   105
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label Label78 
         BackStyle       =   0  'Transparent
         Caption         =   "Fare : Rs 100000"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11880
         TabIndex        =   104
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label77 
         BackStyle       =   0  'Transparent
         Caption         =   "( First Class )"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11880
         TabIndex        =   103
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label59 
         BackColor       =   &H80000005&
         Caption         =   "243 432 234"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9840
         TabIndex        =   95
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000005&
         Caption         =   "Date of Journey"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   85
         Top             =   6480
         Width           =   1815
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000005&
         Caption         =   "Booking Status"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   29
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000005&
         Caption         =   "243 432 234"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9840
         TabIndex        =   28
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000005&
         Caption         =   "Availabe"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8760
         TabIndex        =   27
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000005&
         Caption         =   "(optional)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   26
         Top             =   6480
         Width           =   735
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   25
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "MOBILE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   24
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "GENDER :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   23
         Top             =   7080
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "PASSPORT ID :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   22
         Top             =   5760
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   21
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "FULL NAME:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   20
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Travell Class"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11880
         TabIndex        =   19
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Available airways"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   18
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Suggesstion"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "DESTINATION:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "FROM:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   3720
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   7695
      Left            =   0
      TabIndex        =   87
      Top             =   0
      Visible         =   0   'False
      Width           =   14055
      Begin VB.CommandButton Command6 
         Height          =   1095
         Left            =   9360
         Picture         =   "user_panel.frx":41D24
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   6120
         Width           =   3015
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2340
         Left            =   6960
         TabIndex        =   91
         Top             =   1200
         Width           =   5415
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   90
         Top             =   720
         Width           =   5415
      End
      Begin VB.PictureBox Picture5 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4005
         Left            =   1920
         Picture         =   "user_panel.frx":42CD9
         ScaleHeight     =   4005
         ScaleWidth      =   4125
         TabIndex        =   88
         Top             =   360
         Width           =   4125
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "gdysg dwvw brtb rbrb tnt rvgbrgb  sydgugcuguygsdcyu ygyuegfueg ufc egfwegdq"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   93
         Top             =   5520
         Width           =   12255
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "gdysg dwvw brtb rbrb tnt rvgbrgb  sydgugcuguygsdcyu ygyuegfueg ufc egfwegdq"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   92
         Top             =   4920
         Width           =   12255
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Customer name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8520
         TabIndex        =   89
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   7695
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Visible         =   0   'False
      Width           =   14055
      Begin VB.CommandButton Command3 
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   78
         Top             =   6600
         Width           =   1695
      End
      Begin VB.TextBox Text14 
         Height          =   1935
         Left            =   360
         TabIndex        =   77
         Top             =   4320
         Width           =   6735
      End
      Begin VB.TextBox Text13 
         Height          =   495
         Left            =   2640
         TabIndex        =   76
         Top             =   2880
         Width           =   4455
      End
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   2640
         TabIndex        =   75
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   2640
         TabIndex        =   74
         Top             =   1440
         Width           =   4455
      End
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4500
         Left            =   8760
         Picture         =   "user_panel.frx":47372
         ScaleHeight     =   4500
         ScaleWidth      =   4500
         TabIndex        =   73
         Top             =   1560
         Width           =   4500
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "Message :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   82
         Top             =   3840
         Width           =   2775
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Number :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   81
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Mail : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   80
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "Complainer's Name : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   79
         Top             =   1440
         Width           =   2775
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   14055
      Begin VB.CommandButton Command7 
         Caption         =   "NEXT"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   97
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List6 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   9480
         TabIndex        =   96
         Top             =   720
         Width           =   1935
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000005&
         Height          =   3375
         Left            =   6360
         TabIndex        =   52
         Top             =   3960
         Width           =   7335
         Begin VB.CommandButton Command2 
            Height          =   855
            Left            =   4800
            Picture         =   "user_panel.frx":518FA
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   2040
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   375
            Left            =   1920
            TabIndex        =   54
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   107282433
            CurrentDate     =   41388
         End
         Begin VB.Label Label53 
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3600
            TabIndex        =   102
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label Label62 
            BackStyle       =   0  'Transparent
            Caption         =   "Label62"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   101
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label38 
            BackColor       =   &H80000005&
            Caption         =   "Date of Journey :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   69
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label39 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Status :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   68
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label40 
            BackColor       =   &H80000005&
            Caption         =   "No seat Available"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   67
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label41 
            BackColor       =   &H80000005&
            Caption         =   "First Class"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   66
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "99"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   65
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label43 
            BackColor       =   &H80000005&
            Caption         =   "out of"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   64
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label Label44 
            BackColor       =   &H80000005&
            Caption         =   "90"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   63
            Top             =   2640
            Width           =   375
         End
         Begin VB.Label Label45 
            BackColor       =   &H80000005&
            Caption         =   "Business Class"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   62
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label46 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "99"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1560
            TabIndex        =   61
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label47 
            BackColor       =   &H80000005&
            Caption         =   "out of"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   60
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label Label48 
            BackColor       =   &H80000005&
            Caption         =   "90"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   59
            Top             =   2640
            Width           =   375
         End
         Begin VB.Label Label49 
            BackColor       =   &H80000005&
            Caption         =   "Economy Class"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   58
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label50 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "99"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3000
            TabIndex        =   57
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label51 
            BackColor       =   &H80000005&
            Caption         =   "out of"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   56
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label Label52 
            BackColor       =   &H80000005&
            Caption         =   "90"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   55
            Top             =   2640
            Width           =   375
         End
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   7680
         TabIndex        =   51
         Top             =   2640
         Width           =   6015
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   50
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   49
         Top             =   720
         Width           =   2295
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   6000
         Left            =   120
         Picture         =   "user_panel.frx":52985
         ScaleHeight     =   6000
         ScaleWidth      =   6000
         TabIndex        =   99
         Top             =   1560
         Width           =   6000
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "choose airlines"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   100
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Label Label60 
         BackStyle       =   0  'Transparent
         Caption         =   "choose airlines"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   98
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "FROM :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   71
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "TO :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   70
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7695
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   14055
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   4080
         Width           =   3735
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         Left            =   240
         TabIndex        =   45
         Top             =   4560
         Width           =   3735
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H80000005&
         Caption         =   "Customer Details"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   4200
         TabIndex        =   32
         Top             =   4080
         Width           =   9735
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri Light"
               Size            =   9.75
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   33
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri Light"
               Size            =   9.75
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   5880
            TabIndex        =   44
            Top             =   1800
            Width           =   3615
         End
         Begin VB.Label Label32 
            BackColor       =   &H80000005&
            Caption         =   "Booking Details"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   43
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri Light"
               Size            =   9.75
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   42
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri Light"
               Size            =   9.75
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7560
            TabIndex        =   40
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label26 
            BackColor       =   &H80000005&
            Caption         =   "Date of Journey :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5880
            TabIndex        =   39
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri Light"
               Size            =   9.75
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   38
            Top             =   2040
            Width           =   3735
         End
         Begin VB.Label Label24 
            BackColor       =   &H80000005&
            Caption         =   "Contact Info :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri Light"
               Size            =   9.75
               Charset         =   0
               Weight          =   300
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1560
            TabIndex        =   36
            Top             =   960
            Width           =   3735
         End
         Begin VB.Label Label22 
            BackColor       =   &H80000005&
            Caption         =   "Address :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label20 
            BackColor       =   &H80000005&
            Caption         =   "Customer Name : "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   3495
         Left            =   0
         Picture         =   "user_panel.frx":581F6
         ScaleHeight     =   3435
         ScaleWidth      =   13995
         TabIndex        =   31
         Top             =   0
         Width           =   14055
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000005&
         Caption         =   "Enter Customer Name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   3720
         Width           =   3855
      End
   End
   Begin VB.Menu Home 
      Caption         =   "Home"
      Index           =   0
   End
   Begin VB.Menu Booking 
      Caption         =   "Booking"
      Index           =   1
   End
   Begin VB.Menu Cancellation 
      Caption         =   "Cancellation"
   End
   Begin VB.Menu Patron 
      Caption         =   "Patron"
      Index           =   3
   End
   Begin VB.Menu Query 
      Caption         =   "Query"
      Index           =   2
   End
   Begin VB.Menu c 
      Caption         =   "Submit Tokken"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conn As ADODB.Connection
Public rs, es, ds, fs As ADODB.Recordset
Dim ch, fid, tw As Integer
Dim dt1 As String

Private Sub Form_QueryUnload(cancel As Integer, unloadmode As Integer)

MsgBox "You have been logged out successfully!"
Form1.Enabled = True
Form1.Show
Unload Me
End Sub




Private Sub Booking_Click(Index As Integer)

    List1.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = False

    Set conn = New ADODB.Connection
    Set fs = New ADODB.Recordset
    conn.Open Form1.cs
    fs.ActiveConnection = conn
    fs.CursorLocation = adUseClient
    fs.CursorType = adOpenDynamic
    fs.LockType = adLockOptimistic
    
    fs.Source = "SELECT * FROM Flights"
    fs.Open

Frame7.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame2.Visible = False
Frame1.Visible = True

List2.Enabled = False
DTPicker2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Combo1.Enabled = False
Command1.Enabled = False

Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
            

DTPicker2.Format = dtpCustom
DTPicker2.CustomFormat = "dd-MM-yyyy"
DTPicker2.Value = Format(Date, "dd-MM-yyyy")

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Label78.Caption = ""
Label82.Caption = ""
Label80.Caption = ""

List1.Clear
List2.Clear

Label17.Caption = ""
Label59.Caption = ""

rs.Source = "SELECT * FROM Location WHERE LOC_DIST LIKE 0"
rs.Open

If rs.EOF <> True Then
rs.MoveFirst
End If

Text1.Text = rs("LOC_NAME")
Text2.SetFocus
rs.Close
tw = 2
End Sub

Private Sub c_Click()
Frame2.Visible = False
Frame4.Visible = False
Frame7.Visible = False
Frame5.Visible = False
Frame1.Visible = False
Frame3.Visible = False
Frame6.Visible = True
End Sub

Private Sub Cancellation_Click()
Frame7.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame2.Visible = False
Frame1.Visible = False
Frame3.Visible = True


    Set conn = New ADODB.Connection
    Set ds = New ADODB.Recordset
    conn.Open Form1.cs
    ds.ActiveConnection = conn
    ds.CursorLocation = adUseClient
    ds.CursorType = adOpenDynamic
    ds.LockType = adLockOptimistic
    
    ds.Source = "SELECT * FROM reservation"
    ds.Open
    
    List5.Clear
    Text15.Text = ""
    
    Label19.Caption = "****************************************************"
    Label58.Caption = "Enter name in the above text to get possible matches"



    If ds.EOF <> True Then
            ds.MoveFirst
        End If
        
        Dim j As Integer
        j = 0
        While ds.EOF <> True
            List5.AddItem (ds("pname") + " - " + Str(ds("ID")))
            List5.ItemData(j) = ds("ID")
            j = j + 1
            ds.MoveNext
        Wend
        

End Sub

Private Sub Command1_Click()

    If Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
        MsgBox "Fields empty!"
        Exit Sub
    End If
    
        Set conn = New ADODB.Connection
        Set es = New ADODB.Recordset
        conn.Open Form1.cs
        es.ActiveConnection = conn
        es.CursorLocation = adUseClient
        es.CursorType = adOpenDynamic
        es.LockType = adLockOptimistic
        
        es.Source = "SELECT * FROM reservation"
        es.Open
        
        es.AddNew
        
        es("FIGHT_ID") = fid
        es("pname") = Text3.Text
        es("paddr") = Text4.Text
        es("pid") = Text5.Text
        es("mob") = Text6.Text
        es("email") = Text7.Text
        
        If Option1.Value = True Then
        es("gender") = "M"
        End If
        If Option2.Value = True Then
        es("gender") = "F"
        End If
        If Option3.Value = True Then
        es("gender") = "O"
        End If
        
        Dim fair As Integer
        
        If Combo1.Text = "First Class" Then
        es("tcl") = 1
        fair = fs("FRA")
        End If
        If Combo1.Text = "Bussiness Class" Then
        es("tcl") = 2
        fair = fs("BRA")
        End If
        If Combo1.Text = "Economy Class" Then
        es("tcl") = 3
        fair = fs("ERA")
        End If
        
        es("dod") = dt1
        
        es("FROM") = Text1.Text
        es("TO") = Text2.Text
        
        rs.Source = "SELECT * FROM Location WHERE LOC_NAME = '" + Text1.Text + "' OR LOC_NAME = '" + Text2.Text + "'"
        rs.Open
        
        Dim dist1, dist2, abdist As Integer
        
        If rs.EOF <> True Then
        rs.MoveFirst
        dist1 = rs("LOC_DIST")
        rs.MoveNext
        dist2 = rs("LOC_DIST")
        rs.Close
        End If
        
        abdist = Abs(dist1 - dist2)
        
        fair = fair * abdist
        If MsgBox(" Confirm Booking? ", vbQuestion + vbYesNo, "Confirm Dialouge") = vbYes Then
                es.Update
                es.Close
                MsgBox " Your Ticket Booked Successfully"
                Booking_Click (0)
        Else
                Exit Sub
        End If
        
End Sub

Private Sub Command2_Click()

    
        Booking_Click (0)
    
    
        Label17.Caption = " " + Label42.Caption + " " + Label46.Caption + " " + Label50.Caption + " "
        Label59.Caption = " " + Label44.Caption + " " + Label48.Caption + " " + Label52.Caption + " "
        
        Text1.Text = Text9.Text
        Text2.Text = Text10.Text
        List2.Enabled = False
        List1.Enabled = False
        
        
        DTPicker2.Value = DTPicker3.Value
        
        
    
    
    Command4.Enabled = False
    Command5.Enabled = False
    
        Text3.Enabled = True
        Text4.Enabled = True
        Text5.Enabled = True
        Text6.Enabled = True
        Text7.Enabled = True
        
        Option1.Enabled = True
        Option2.Enabled = True
        Option3.Enabled = True
        
        Combo1.Enabled = True
        Command1.Enabled = True
        
    
    
    Frame7.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
    Frame5.Visible = False
    Frame6.Visible = False
    Frame2.Visible = False
    Frame1.Visible = True
    
        

End Sub

Private Sub Command3_Click()
Set conn = New ADODB.Connection
    conn.Open Form1.cs
Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    
    
    If Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Then
    MsgBox "Fields Empty! Please fill up!!"
    Exit Sub
    End If
    
    rs.Source = "SELECT * FROM tickets"
    rs.Open
    rs.AddNew
         rs("cname") = Text11.Text
         rs("cmail") = Text12.Text
         rs("cnum") = Text13.Text
         rs("mess") = Text14.Text
    rs.Update
    rs.Close
        MsgBox "  Your Query submitted successfully ! We will contact you shortly"
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        Text14.Text = ""

End Sub

Private Sub Command4_Click()
    Dim tu, frm  As String
    
        Set conn = New ADODB.Connection
        conn.Open Form1.cs
        Set rs = New ADODB.Recordset
        rs.ActiveConnection = conn
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenDynamic
        rs.LockType = adLockOptimistic
        
        List2.Clear
        tu = Text2.Text
        frm = Text1.Text
        
        rs.Source = "SELECT * FROM Flights WHERE FROM = '" & frm & "' AND TO = '" & tu & "'"
        rs.Open
        If rs.EOF <> True Then
            rs.MoveFirst
        End If
        Dim j As Integer
        j = 0
        While rs.EOF <> True
            List2.AddItem (rs("AIRLINES") + " - " + Str(rs("FIGHT_ID")))
            List2.ItemData(j) = rs("FIGHT_ID")
            j = j + 1
            rs.MoveNext
        Wend
        
        
        rs.Close
        
        rs.Source = "SELECT * FROM Flights WHERE FROM = '" & frm & "' AND IL LIKE '%" & tu & "%'"
        rs.Open
        If rs.EOF <> True Then
            rs.MoveFirst
            
        End If
        
        While rs.EOF <> True
            List2.AddItem (rs("AIRLINES") + " - " + Str(rs("FIGHT_ID")))
            List2.ItemData(j) = rs("FIGHT_ID")
            j = j + 1
            rs.MoveNext
        Wend
        
        
        rs.Close
        
        
        rs.Source = "SELECT * FROM Flights WHERE TO = '" & frm & "' AND IL LIKE '%" & tu & "%'"
        rs.Open
        If rs.EOF <> True Then
            rs.MoveFirst
            
        End If
        
        While rs.EOF <> True
            List2.AddItem (rs("AIRLINES") + " - " + Str(rs("FIGHT_ID")))
            List2.ItemData(j) = rs("FIGHT_ID")
            j = j + 1
            rs.MoveNext
        Wend
        
        
        rs.Close
        
        rs.Source = "SELECT * FROM Flights WHERE IL LIKE '%" & frm & "%' AND IL LIKE '%" & tu & "%'"
        rs.Open
        If rs.EOF <> True Then
            rs.MoveFirst
            
        End If
        
        While rs.EOF <> True
            List2.AddItem (rs("AIRLINES") + " - " + Str(rs("FIGHT_ID")))
             List2.ItemData(j) = rs("FIGHT_ID")
            j = j + 1
            rs.MoveNext
        Wend
        
        
        rs.Close
        
        List2.Enabled = True
End Sub

Private Sub Command5_Click()

Dim dtbool As Boolean
Dim dt, dth() As String
dt = "0"

If fs("mon") Then
dt = dt + "_2"
End If

If fs("tue") Then
dt = dt + "_3"
End If

If fs("wed") Then
dt = dt + "_4"
End If

If fs("thur") Then
dt = dt + "_5"
End If

If fs("fri") Then
dt = dt + "_6"
End If

If fs("sat") Then
dt = dt + "_7"
End If

If fs("sun") Then
dt = dt + "_1"
End If

dth = Split(dt, "_")

Dim dtd, dts, k As Integer
dtd = DTPicker2.DayOfWeek

dts = UBound(dth)
k = 0
    dtbool = False
    
    
    For k = 0 To dts
    
        If dtd = Val(dth(k)) Then
            dtbool = True
            Exit For
        End If
   Next
    
    If (Not (dtbool)) Then
        MsgBox "The selected flight doesnt run on this day!"
            
            Text3.Enabled = False
            Text4.Enabled = False
            Text5.Enabled = False
            Text6.Enabled = False
            Text7.Enabled = False
            
            Option1.Enabled = False
            Option2.Enabled = False
            Option3.Enabled = False
            
            Combo1.Enabled = False
            Command1.Enabled = False
        
        Exit Sub
    End If
    
    'DTPicker2.Format = "dd-mm-yyyy"
    
    DTPicker2.Format = dtpCustom
    DTPicker2.CustomFormat = "dd-MM-yyyy"
    dt1 = Format(DTPicker2.Value, "dd-MM-yyyy")
    
    
    es.Source = "SELECT * FROM reservation WHERE FIGHT_ID=" + Str(fs("FIGHT_ID")) + " AND dod LIKE '%" + dt1 + "%'"
    es.Open
    
    
    Dim cB, cF, cE As Boolean
    cB = False
    cF = False
    cE = False
    
    Dim fc, bc, ec As Integer
    es.Filter = "tcl = 1"
    fc = fs("FMA") - es.RecordCount
    
    If fc > 0 Then
        cF = True
    End If
    
    es.Filter = "tcl = 2"
    bc = fs("BMA") - es.RecordCount
   
    If bc > 0 Then
        cB = True
    End If
    
    es.Filter = "tcl = 3"
    ec = fs("EMA") - es.RecordCount
    
    If ec > 0 Then
        cE = True
    End If
    
    es.Close
    
    Label59.Caption = " " + Str(fs("FMA")) + " " + Str(fs("BMA")) + " " + Str(fs("EMA")) + " "
    
    If Not (cB Or cE Or cF) Then
        Label17.ForeColor = vbRed
        Label17.Caption = "No seat Availabe"
        Exit Sub
    Else
        Label17.ForeColor = vbGreen
        Label17.Caption = " " + Str(fc) + " " + Str(bc) + " " + Str(ec) + " "
        
        
        Text3.Enabled = True
        Text4.Enabled = True
        Text5.Enabled = True
        Text6.Enabled = True
        Text7.Enabled = True
        
        Option1.Enabled = True
        Option2.Enabled = True
        Option3.Enabled = True
        
        Combo1.Enabled = True
        Command1.Enabled = True
        
        Combo1.Clear
        
        If cF Then
            Combo1.AddItem ("First Class")
        End If
        
        If cB Then
            Combo1.AddItem ("Bussiness Class")
        End If
        
        If cE Then
            Combo1.AddItem ("Economy Class")
        End If
        
    End If
    
    Dim dist1, dist2, abdist As Integer
        
        rs.Source = "SELECT * FROM Location Where LOC_NAME = '" + Text1.Text + "' OR '" + Text2.Text + "'"
        rs.Open
        
        If rs.EOF <> True Then
        rs.MoveFirst
        dist1 = rs("LOC_DIST")
        rs.MoveNext
        dist2 = rs("LOC_DIST")
        rs.Close
        End If
        
        abdist = Abs(dist1 - dist2)
        
        Label78.Caption = "Fare : Rs " + Str(Val(fs("FRA")) * abdist)
        Label82.Caption = "Fare : Rs " + Str(Val(fs("BRA")) * abdist)
        Label80.Caption = "Fare : Rs " + Str(Val(fs("ERA")) * abdist)
        
        
        
    
    
    
End Sub

Private Sub Command6_Click()
    If Not (List5.ListIndex >= 0) Then
        MsgBox "Please enter/ choose a valid name"
        Exit Sub
    End If
    
    If MsgBox("Delete " & ds("pname") & " from reservation chart? Action cannot be undone. Continue??", vbQuestion + vbYesNo, "Delete passenger") = vbYes Then
        ds.Delete
        MsgBox "Deleted, Booking Cleared!"
        Cancellation_Click
    Else
        Exit Sub
    End If
End Sub

Private Sub Command7_Click()
        Dim tu, frm  As String
    
        Set conn = New ADODB.Connection
        conn.Open Form1.cs
        Set rs = New ADODB.Recordset
        rs.ActiveConnection = conn
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenDynamic
        rs.LockType = adLockOptimistic
        
        List4.Clear
        tu = Text10.Text
        frm = Text9.Text
        
        Dim ff As Boolean
        ff = False
        rs.Source = "SELECT * FROM Flights WHERE FROM = '" & frm & "' AND TO = '" & tu & "'"
        rs.Open
        If rs.EOF <> True Then
            rs.MoveFirst
            ff = True
        End If
        Dim j As Integer
        j = 0
        While rs.EOF <> True
            List4.AddItem (rs("AIRLINES") + " - " + Str(rs("FIGHT_ID")))
            List4.ItemData(j) = rs("FIGHT_ID")
            j = j + 1
            rs.MoveNext
        Wend
        
        
        rs.Close
        
        rs.Source = "SELECT * FROM Flights WHERE FROM = '" & frm & "' AND IL LIKE '%" & tu & "%'"
        rs.Open
        If rs.EOF <> True Then
            rs.MoveFirst
            ff = True
        End If
        
        While rs.EOF <> True
            List4.AddItem (rs("AIRLINES") + " - " + Str(rs("FIGHT_ID")))
            List4.ItemData(j) = rs("FIGHT_ID")
            j = j + 1
            rs.MoveNext
        Wend
        
        
        rs.Close
        
        
        rs.Source = "SELECT * FROM Flights WHERE TO = '" & frm & "' AND IL LIKE '%" & tu & "%'"
        rs.Open
        If rs.EOF <> True Then
            rs.MoveFirst
            ff = True
        End If
        
        While rs.EOF <> True
            List4.AddItem (rs("AIRLINES") + " - " + Str(rs("FIGHT_ID")))
            List4.ItemData(j) = rs("FIGHT_ID")
            j = j + 1
            rs.MoveNext
        Wend
        
        
        rs.Close
        
        rs.Source = "SELECT * FROM Flights WHERE IL LIKE '%" & frm & "%' AND IL LIKE '%" & tu & "%'"
        rs.Open
        If rs.EOF <> True Then
            rs.MoveFirst
            ff = True
            
        End If
        
        While rs.EOF <> True
            List4.AddItem (rs("AIRLINES") + " - " + Str(rs("FIGHT_ID")))
             List4.ItemData(j) = rs("FIGHT_ID")
            j = j + 1
            rs.MoveNext
        Wend
        
        
        rs.Close
        List4.Enabled = True
        Label61.Caption = " Flights from " + frm + " to " + tu
        If Not ff Then
        List4.AddItem ("No such flight found!")
        List4.Enabled = False
        Frame5.Visible = False
        End If
End Sub





Private Sub DTPicker3_Change()
List4_Click
End Sub

Private Sub Form_Load()

Set conn = New ADODB.Connection
    conn.Open Form1.cs
Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic


Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame1.Visible = False
Frame2.Visible = False
End Sub

Private Sub Home_Click(Index As Integer)
Frame7.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame1.Visible = False
Frame2.Visible = False
End Sub

Private Sub Label28_Click()
If tw = 1 Then
Text1.Text = "Mumbai"
Else
Text2.Text = "Mumbai"
End If
End Sub

Private Sub Label29_Click()
If tw = 1 Then
Text1.Text = "Delhi"
Else
Text2.Text = "Delhi"
End If
End Sub

Private Sub Label34_Click()
If tw = 1 Then
Text1.Text = "Kolkata"
Else
Text2.Text = "Kolkata"
End If
End Sub

Private Sub Label35_Click()
If tw = 1 Then
Text1.Text = "Patna"
Else
Text2.Text = "Patna"
End If
End Sub

Private Sub Label63_Click()
If tw = 1 Then
Text1.Text = "Bangalore"
Else
Text2.Text = "Bangalore"
End If
End Sub

Private Sub Label64_Click()
If tw = 1 Then
Text1.Text = "Sydney"
Else
Text2.Text = "Sydney"
End If
End Sub

Private Sub Label65_Click()
If tw = 1 Then
Text1.Text = "San Fransisco"
Else
Text2.Text = "San Fransisco"
End If
End Sub

Private Sub Label66_Click()
If tw = 1 Then
Text1.Text = "Austin"
Else
Text2.Text = "Austin"
End If
End Sub

Private Sub Label67_Click()
If tw = 1 Then
Text1.Text = "London"
Else
Text2.Text = "London"
End If
End Sub

Private Sub Label68_Click()
If tw = 1 Then
Text1.Text = "Berlin"
Else
Text2.Text = "Berlin"
End If
End Sub

Private Sub Label69_Click()
If tw = 1 Then
Text1.Text = "Dubai"
Else
Text2.Text = "Dubai"
End If
End Sub

Private Sub Label70_Click()
If tw = 1 Then
Text1.Text = "Hong Kong"
Else
Text2.Text = "Hong Kong"
End If
End Sub

Private Sub Label71_Click()
If tw = 1 Then
Text1.Text = "Beijing"
Else
Text2.Text = "Beijing"
End If
End Sub

Private Sub Label72_Click()
If tw = 1 Then
Text1.Text = "Tokyo"
Else
Text2.Text = "Tokyo"
End If
End Sub

Private Sub Label73_Click()
If tw = 1 Then
Text1.Text = "New York City"
Else
Text2.Text = "New York City"
End If
End Sub

Private Sub Label74_Click()
If tw = 1 Then
Text1.Text = "Capetown"
Else
Text2.Text = "Capetown"
End If
End Sub

Private Sub Label75_Click()
If tw = 1 Then
Text1.Text = "Singapore"
Else
Text2.Text = "Singapore"
End If
End Sub

Private Sub Label76_Click()
If tw = 1 Then
Text1.Text = "Auckland"
Else
Text2.Text = "Auckland"
End If
End Sub

Private Sub List1_Click()
If List1.ListIndex >= 0 Then
    If ch = 1 Then
        Text1.Text = List1.List(List1.ListIndex)
    Else
        If ch = 2 Then
        Text2.Text = List1.List(List1.ListIndex)
        End If
    End If
End If
End Sub

Private Sub List2_Click()
fs.Filter = "FIGHT_ID LIKE " + Str(List2.ItemData(List2.ListIndex))
fid = fs("FIGHT_ID")
DTPicker2.Enabled = True
Command5.Enabled = True

End Sub

Private Sub List3_Click()
    rs.Source = "SELECT * FROM reservation WHERE ID = " + Str(List3.ItemData(List3.ListIndex))
    rs.Open
    
        If rs.EOF <> True Then
            Label21.Caption = rs("pname")
            Label23.Caption = rs("paddr")
            Label25.Caption = rs("mob") + ", " + rs("email")
            Label27.Caption = rs("dod")
            Label33.Caption = "Flight from " + rs("FROM") + " to " + rs("TO")
            If rs("gender") = "M" Then
                Label31.Caption = "MALE"
            Else
                If rs("gender") = "F" Then
                Label31.Caption = "FEMALE"
                Else
                Label31.Caption = "Other"
                End If
            End If
        
        
        
        End If
    rs.Close
End Sub

Private Sub List4_Click()
   
    Set conn = New ADODB.Connection
    Set es = New ADODB.Recordset
    conn.Open Form1.cs
    es.ActiveConnection = conn
    es.CursorLocation = adUseClient
    es.CursorType = adOpenDynamic
    es.LockType = adLockOptimistic
   
   
   Frame5.Visible = True
   rs.Source = "SELECT * FROM Flights WHERE FIGHT_ID = " + Str(List4.ItemData(List4.ListIndex))
   rs.Open

Dim dtbool As Boolean
Dim dt, LB, dth() As String
dt = "0"
LB = "Running On "

If rs("mon") Then
dt = dt + "_2"
LB = LB + " Monday"
End If

If rs("tue") Then
dt = dt + "_3"
LB = LB + " Tuesday"
End If

If rs("wed") Then
dt = dt + "_4"
LB = LB + " Wednesday"
End If

If rs("thur") Then
dt = dt + "_5"
LB = LB + " Thursday"
End If

If rs("fri") Then
dt = dt + "_6"
LB = LB + " Friday"
End If

If rs("sat") Then
dt = dt + "_7"
LB = LB + " Saturday"
End If

If rs("sun") Then
dt = dt + "_1"
LB = LB + " Sunday"
End If

dth = Split(dt, "_")
fid = rs("FIGHT_ID")
        
        
        Dim dtd, dts, k As Integer
        dtd = DTPicker3.DayOfWeek

dts = UBound(dth)
k = 0
    dtbool = False
    
    
    For k = 0 To dts
    
        If dtd = Val(dth(k)) Then
            dtbool = True
            Exit For
        End If
    Next
    
    If (Not (dtbool)) Then
        MsgBox "The selected flight doesnt run on this day!"
        rs.Close
        
        Label40.Caption = ""
        
        Label42.Caption = ""
        Label46.Caption = ""
        Label50.Caption = ""
        
        
        Label44.Caption = ""
        Label48.Caption = ""
        Label52.Caption = ""
        Label53.Caption = ""
        Label62.Caption = ""
        Command2.Enabled = False
        
        Combo1.Clear
        
        
        
        Exit Sub
    End If
    
    DTPicker3.Format = dtpCustom
    DTPicker3.CustomFormat = "dd-MM-yyyy"
    dt1 = Format(DTPicker3.Value, "dd-MM-yyyy")
    
If List4.ListIndex >= 0 Then
    es.Source = "SELECT * FROM reservation WHERE FIGHT_ID=" + Str(List4.ItemData(List4.ListIndex)) + " AND dod LIKE '%" + dt1 + "%'"
   
    es.Open
     
    
    Dim cB, cF, cE As Boolean
    cB = False
    cF = False
    cE = False
    
    Dim fc, bc, ec As Integer
    es.Filter = "tcl = 1"
    fc = rs("FMA") - es.RecordCount
    
    
    If fc > 0 Then
        cF = True
    End If
    
    es.Filter = "tcl = 2"
    bc = rs("BMA") - es.RecordCount
   
    If bc > 0 Then
        cB = True
    End If
    
    es.Filter = "tcl = 3"
    ec = rs("EMA") - es.RecordCount
   
    If ec > 0 Then
        cE = True
    End If
    
    es.Close
    
    
    If Not (cB Or cE Or cF) Then
        Label40.ForeColor = vbRed
        Label40.Caption = "No seat Availabe"
        Exit Sub
    Else
        Label40.ForeColor = vbGreen
        Label40.Caption = "Available"
        
        Label42.Caption = Str(fc)
        Label46.Caption = Str(bc)
        Label50.Caption = Str(ec)
        
       
        Label44.Caption = rs("FMA")
        Label48.Caption = rs("BMA")
        Label52.Caption = rs("EMA")
        Label53.Caption = LB
        Label62.Caption = "Dept on " + rs("dept")
        Command2.Enabled = True
        
        Combo1.Clear
        
        If cF Then
            Combo1.AddItem ("First Class")
        End If
        
        If cB Then
            Combo1.AddItem ("Bussiness Class")
        End If
        
        If cE Then
            Combo1.AddItem ("Economy Class")
        End If
        
        
    End If

End If

rs.Close
End Sub

Private Sub List5_Click()

ds.Filter = "ID LIKE " + Str(List5.ItemData(List5.ListIndex))
Label19.Caption = "Name : " + ds("pname") + ", contact : " + ds("mob")
Label58.Caption = "Booking from " + ds("FROM") + " to " + ds("TO") + " on " + ds("dod")
End Sub

Private Sub List6_Click()
If List6.ListIndex >= 0 Then
    If ch = 1 Then
        Text9.Text = List6.List(List6.ListIndex)
    Else
        If ch = 2 Then
        Text10.Text = List6.List(List6.ListIndex)
        End If
    End If
End If

End Sub

Private Sub Patron_Click(Index As Integer)

        rs.Source = "SELECT * FROM reservation "
        
        
        rs.Open
        If rs.EOF <> True Then
            rs.MoveFirst
        End If
        List3.Clear
        Dim j As Integer
        j = 0
        While rs.EOF <> True
            List3.AddItem (rs("pname") + " - " + Str(rs("ID")))
            List3.ItemData(j) = rs("ID")
            j = j + 1
            rs.MoveNext
        Wend
        
        
        rs.Close





Frame4.Visible = False
Frame3.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame1.Visible = False
Frame2.Visible = True
Frame7.Visible = True
End Sub

Private Sub Query_Click(Index As Integer)
Frame2.Visible = False
Frame6.Visible = False
Frame3.Visible = False
Frame1.Visible = False
Frame7.Visible = False
Frame4.Visible = True
Frame5.Visible = False


            Label42.Caption = ""
            Label46.Caption = ""
            Label50.Caption = ""
            
            Label44.Caption = ""
            Label48.Caption = ""
            Label52.Caption = ""
            
            Label40.Caption = ""
            Command2.Enabled = False

Text9.Text = ""
Text10.Text = ""
Label61.Caption = ""
List6.Clear
            
        rs.Source = "SELECT * FROM Flights"
        rs.Open
        If rs.EOF <> True Then
            rs.MoveFirst
        End If
        
        While rs.EOF <> True
            List4.AddItem (rs("AIRLINES") + " - " + Str(rs("FIGHT_ID")))
             List4.ItemData(j) = rs("FIGHT_ID")
            j = j + 1
            rs.MoveNext
        Wend
        
        
        rs.Close

    
End Sub

Private Sub Text1_Change()
    ch = 1
    
    Set conn = New ADODB.Connection
    Set es = New ADODB.Recordset
    conn.Open Form1.cs
    es.ActiveConnection = conn
    es.CursorLocation = adUseClient
    es.CursorType = adOpenDynamic
    es.LockType = adLockOptimistic
    
    es.Source = "SELECT * FROM Location WHERE LOC_NAME LIKE '" & Text1.Text & "%'"
    
    es.Open
      
        Dim num, i As Integer
        num = 0
        i = 0
        If es.EOF <> True Then
        num = es.RecordCount
        Else
        MsgBox "No such destination"
        End If
        
        List1.Clear
            While i < num
                List1.AddItem (es("LOC_NAME"))
                i = i + 1
                es.MoveNext
            Wend
    es.Close

    

End Sub

Private Sub Text1_GotFocus()
tw = 1
End Sub

Private Sub Text10_Change()
  ch = 2
    
    Set conn = New ADODB.Connection
    Set es = New ADODB.Recordset
    conn.Open Form1.cs
    es.ActiveConnection = conn
    es.CursorLocation = adUseClient
    es.CursorType = adOpenDynamic
    es.LockType = adLockOptimistic
    
    es.Source = "SELECT * FROM Location WHERE LOC_NAME LIKE '" & Text10.Text & "%'"
    
    es.Open
      
        Dim num, i As Integer
        num = 0
        i = 0
        If es.EOF <> True Then
        num = es.RecordCount
        Else
        MsgBox "No such destination"
        End If
        
        List6.Clear
            While i < num
                List6.AddItem (es("LOC_NAME"))
                i = i + 1
                es.MoveNext
            Wend
    es.Close

    

End Sub

Private Sub Text15_Change()
        Set conn = New ADODB.Connection
        conn.Open Form1.cs
        Set rs = New ADODB.Recordset
        rs.ActiveConnection = conn
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenDynamic
        rs.LockType = adLockOptimistic
        
        List5.Clear
        
        rs.Source = "SELECT * FROM reservation WHERE pname LIKE '" + Text15.Text + "%'"
        
        
        rs.Open
        If rs.EOF <> True Then
            rs.MoveFirst
        End If
        
        Dim j As Integer
        j = 0
        While rs.EOF <> True
            List5.AddItem (rs("pname") + " - " + Str(rs("ID")))
            List5.ItemData(j) = rs("ID")
            j = j + 1
            rs.MoveNext
        Wend
        
        
        rs.Close
End Sub

Private Sub Text2_Change()
    ch = 2
    
    Set conn = New ADODB.Connection
    Set es = New ADODB.Recordset
    conn.Open Form1.cs
    es.ActiveConnection = conn
    es.CursorLocation = adUseClient
    es.CursorType = adOpenDynamic
    es.LockType = adLockOptimistic
    
    es.Source = "SELECT * FROM Location WHERE LOC_NAME LIKE '" & Text2.Text & "%'"
    
    es.Open
      
        Dim num, i As Integer
        num = 0
        i = 0
        If es.EOF <> True Then
        num = es.RecordCount
        Else
        MsgBox "No such destination"
        End If
        
        List1.Clear
            While i < num
                List1.AddItem (es("LOC_NAME"))
                i = i + 1
                es.MoveNext
            Wend
    es.Close

    

End Sub

Private Sub Text2_GotFocus()
tw = 2
End Sub

Private Sub Text8_Change()

        Set conn = New ADODB.Connection
        conn.Open Form1.cs
        Set rs = New ADODB.Recordset
        rs.ActiveConnection = conn
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenDynamic
        rs.LockType = adLockOptimistic
        
        List3.Clear
        
        rs.Source = "SELECT * FROM reservation WHERE pname LIKE '" + Text8.Text + "%'"
        
        
        rs.Open
        If rs.EOF <> True Then
            rs.MoveFirst
        End If
        
        Dim j As Integer
        j = 0
        While rs.EOF <> True
            List3.AddItem (rs("pname") + " - " + Str(rs("ID")))
            List3.ItemData(j) = rs("ID")
            j = j + 1
            rs.MoveNext
        Wend
        
        
        rs.Close

End Sub

Private Sub Text9_Change()
  ch = 1
    
    Set conn = New ADODB.Connection
    Set es = New ADODB.Recordset
    conn.Open Form1.cs
    es.ActiveConnection = conn
    es.CursorLocation = adUseClient
    es.CursorType = adOpenDynamic
    es.LockType = adLockOptimistic
    
    es.Source = "SELECT * FROM Location WHERE LOC_NAME LIKE '" & Text9.Text & "%'"
    
    es.Open
      
        Dim num, i As Integer
        num = 0
        i = 0
        If es.EOF <> True Then
        num = es.RecordCount
        Else
        MsgBox "No such destination"
        End If
        
        List6.Clear
            While i < num
                List6.AddItem (es("LOC_NAME"))
                i = i + 1
                es.MoveNext
            Wend
    es.Close

    

End Sub
