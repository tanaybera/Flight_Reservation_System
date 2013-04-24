VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   15975
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "user_panel.frx":0000
   ScaleHeight     =   8100
   ScaleWidth      =   15975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   7695
      Left            =   0
      TabIndex        =   87
      Top             =   0
      Width           =   14055
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4500
         Left            =   8760
         Picture         =   "user_panel.frx":19D9C
         ScaleHeight     =   4500
         ScaleWidth      =   4500
         TabIndex        =   88
         Top             =   1560
         Width           =   4500
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "Complainer"
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
         TabIndex        =   89
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7695
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   14055
      Begin VB.Frame Frame4 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Width           =   14055
         Begin VB.Frame Frame5 
            BackColor       =   &H80000005&
            Height          =   3375
            Left            =   6360
            TabIndex        =   68
            Top             =   3960
            Width           =   7335
            Begin VB.CommandButton Command2 
               Height          =   855
               Left            =   4800
               Picture         =   "user_panel.frx":24324
               Style           =   1  'Graphical
               TabIndex        =   86
               Top             =   2040
               Width           =   2295
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   375
               Left            =   1920
               TabIndex        =   70
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
               Format          =   93257729
               CurrentDate     =   41388
            End
            Begin VB.Label Label53 
               BackColor       =   &H80000005&
               Caption         =   "Status"
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
               Left            =   5160
               TabIndex        =   85
               Top             =   1560
               Width           =   975
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
               TabIndex        =   84
               Top             =   2640
               Width           =   375
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
               TabIndex        =   83
               Top             =   2640
               Width           =   615
            End
            Begin VB.Label Label50 
               BackColor       =   &H80000005&
               Caption         =   "99"
               BeginProperty Font 
                  Name            =   "Calibri Light"
                  Size            =   36
                  Charset         =   0
                  Weight          =   300
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   3480
               TabIndex        =   82
               Top             =   1800
               Width           =   975
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
               TabIndex        =   81
               Top             =   1560
               Width           =   1215
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
               TabIndex        =   80
               Top             =   2640
               Width           =   375
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
               TabIndex        =   79
               Top             =   2640
               Width           =   615
            End
            Begin VB.Label Label46 
               BackColor       =   &H80000005&
               Caption         =   "99"
               BeginProperty Font 
                  Name            =   "Calibri Light"
                  Size            =   36
                  Charset         =   0
                  Weight          =   300
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   1920
               TabIndex        =   78
               Top             =   1800
               Width           =   975
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
               TabIndex        =   77
               Top             =   1560
               Width           =   1455
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
               TabIndex        =   76
               Top             =   2640
               Width           =   375
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
               TabIndex        =   75
               Top             =   2640
               Width           =   615
            End
            Begin VB.Label Label42 
               BackColor       =   &H80000005&
               Caption         =   "99"
               BeginProperty Font 
                  Name            =   "Calibri Light"
                  Size            =   36
                  Charset         =   0
                  Weight          =   300
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   480
               TabIndex        =   74
               Top             =   1800
               Width           =   975
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
               TabIndex        =   73
               Top             =   1560
               Width           =   1215
            End
            Begin VB.Label Label40 
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
               Left            =   4800
               TabIndex        =   72
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label39 
               BackColor       =   &H80000005&
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
               Left            =   3840
               TabIndex        =   71
               Top             =   480
               Width           =   855
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
            TabIndex        =   67
            Top             =   2280
            Width           =   5055
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
            Left            =   9360
            TabIndex        =   65
            Text            =   "Text10"
            Top             =   1560
            Width           =   3375
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
            Left            =   9360
            TabIndex        =   64
            Text            =   "Text9"
            Top             =   960
            Width           =   3375
         End
         Begin VB.PictureBox Picture3 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   6000
            Left            =   0
            Picture         =   "user_panel.frx":253AF
            ScaleHeight     =   6000
            ScaleWidth      =   6000
            TabIndex        =   62
            Top             =   1680
            Width           =   6000
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
            Left            =   7680
            TabIndex        =   66
            Top             =   1560
            Width           =   1095
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
            Left            =   7680
            TabIndex        =   63
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
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
         TabIndex        =   43
         Top             =   4080
         Width           =   9735
         Begin VB.VScrollBar VScroll1 
            Height          =   855
            Left            =   5040
            TabIndex        =   57
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label35 
            BackColor       =   &H80000005&
            Caption         =   "Status"
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
            TabIndex        =   60
            Top             =   2640
            Width           =   1695
         End
         Begin VB.Label Label34 
            Caption         =   "Label21"
            Height          =   375
            Left            =   7560
            TabIndex        =   59
            Top             =   2640
            Width           =   1935
         End
         Begin VB.Label Label33 
            Caption         =   "Label33"
            Height          =   615
            Left            =   5880
            TabIndex        =   58
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
            TabIndex        =   56
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label31 
            Caption         =   "Label30"
            Height          =   375
            Left            =   3240
            TabIndex        =   55
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label Label30 
            Caption         =   "Label30"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label Label29 
            Caption         =   "Label21"
            Height          =   375
            Left            =   7560
            TabIndex        =   53
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label28 
            BackColor       =   &H80000005&
            Caption         =   "Return On :"
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
            TabIndex        =   52
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label27 
            Caption         =   "Label21"
            Height          =   375
            Left            =   7560
            TabIndex        =   51
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
            TabIndex        =   50
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label25 
            Caption         =   "Label21"
            Height          =   375
            Left            =   1560
            TabIndex        =   49
            Top             =   1920
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
            TabIndex        =   48
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label23 
            Caption         =   "Label21"
            Height          =   855
            Left            =   1560
            TabIndex        =   47
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
            TabIndex        =   46
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "Label21"
            Height          =   375
            Left            =   1560
            TabIndex        =   45
            Top             =   480
            Width           =   3735
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
            TabIndex        =   44
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.ListBox List3 
         Height          =   2985
         Left            =   240
         TabIndex        =   42
         Top             =   4560
         Width           =   3735
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   4080
         Width           =   3735
      End
      Begin VB.PictureBox Picture2 
         Height          =   3495
         Left            =   0
         Picture         =   "user_panel.frx":2AC20
         ScaleHeight     =   3435
         ScaleWidth      =   13995
         TabIndex        =   39
         Top             =   0
         Width           =   14055
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000005&
         Caption         =   "Enter Ticket Number or Customer Name"
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
         TabIndex        =   40
         Top             =   3720
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "booking"
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   14055
      Begin VB.CommandButton Command1 
         Height          =   855
         Left            =   11760
         Picture         =   "user_panel.frx":4503F
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   6600
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   7560
         TabIndex        =   31
         Top             =   6720
         Width           =   3135
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   7560
         TabIndex        =   29
         Top             =   6360
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   11760
         TabIndex        =   28
         Top             =   6000
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   93257729
         CurrentDate     =   41382
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   11760
         TabIndex        =   25
         Top             =   5280
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   93257729
         CurrentDate     =   41382
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H80000005&
         Caption         =   "return"
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
         Left            =   12840
         TabIndex        =   24
         Top             =   4680
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H80000005&
         Caption         =   "one way"
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
         Left            =   11760
         TabIndex        =   23
         Top             =   4680
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   11760
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   4200
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
         Left            =   9480
         TabIndex        =   21
         Top             =   7320
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
         Left            =   8520
         TabIndex        =   20
         Top             =   7320
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
         Left            =   7680
         TabIndex        =   19
         Top             =   7320
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   7560
         TabIndex        =   18
         Top             =   6000
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         Height          =   735
         Left            =   7560
         TabIndex        =   17
         Top             =   5160
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   7560
         TabIndex        =   16
         Top             =   4680
         Width           =   3135
      End
      Begin VB.ListBox List2 
         Height          =   3375
         Left            =   3120
         TabIndex        =   7
         Top             =   4200
         Width           =   2295
      End
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   360
         TabIndex        =   6
         Top             =   5640
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   4920
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   4080
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   3480
         Left            =   0
         Negotiate       =   -1  'True
         Picture         =   "user_panel.frx":45CFB
         ScaleHeight     =   14.25
         ScaleMode       =   4  'Character
         ScaleWidth      =   120
         TabIndex        =   1
         Top             =   120
         Width           =   14460
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
         Left            =   6000
         TabIndex        =   36
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000005&
         Height          =   495
         Left            =   9240
         TabIndex        =   35
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000005&
         Height          =   495
         Left            =   7560
         TabIndex        =   34
         Top             =   3960
         Width           =   1575
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
         Left            =   6720
         TabIndex        =   33
         Top             =   6720
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
         Left            =   6000
         TabIndex        =   32
         Top             =   6720
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
         Left            =   6000
         TabIndex        =   30
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000005&
         Caption         =   "Returning On"
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
         Left            =   11760
         TabIndex        =   27
         Top             =   5760
         Width           =   1335
      End
      Begin VB.Label Label11 
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
         Left            =   11760
         TabIndex        =   26
         Top             =   5040
         Width           =   1815
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
         Left            =   6000
         TabIndex        =   15
         Top             =   7320
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
         Left            =   6000
         TabIndex        =   14
         Top             =   6000
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
         Left            =   6000
         TabIndex        =   13
         Top             =   5160
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
         Left            =   6000
         TabIndex        =   12
         Top             =   4680
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
         Left            =   11760
         TabIndex        =   11
         Top             =   3840
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
         Left            =   3120
         TabIndex        =   9
         Top             =   3840
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
         TabIndex        =   8
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
         TabIndex        =   4
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
         TabIndex        =   2
         Top             =   3720
         Width           =   735
      End
   End
   Begin VB.Label Label5 
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
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   735
   End
   Begin VB.Menu Home 
      Caption         =   "Home"
      Index           =   0
   End
   Begin VB.Menu Booking 
      Caption         =   "Booking"
      Index           =   1
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
Private Sub c_Click()

End Sub
