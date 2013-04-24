VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   11430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form4"
   ScaleHeight     =   11430
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
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
         Picture         =   "temp_frame.frx":0000
         ScaleHeight     =   14.25
         ScaleMode       =   4  'Character
         ScaleWidth      =   120
         TabIndex        =   19
         Top             =   120
         Width           =   14460
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   4080
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   4920
         Width           =   2175
      End
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   360
         TabIndex        =   16
         Top             =   5640
         Width           =   2175
      End
      Begin VB.ListBox List2 
         Height          =   3375
         Left            =   3120
         TabIndex        =   15
         Top             =   4200
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   7560
         TabIndex        =   14
         Top             =   4680
         Width           =   3135
      End
      Begin VB.TextBox Text4 
         Height          =   735
         Left            =   7560
         TabIndex        =   13
         Top             =   5160
         Width           =   3135
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   7560
         TabIndex        =   12
         Top             =   6000
         Width           =   3135
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
         TabIndex        =   11
         Top             =   7320
         Width           =   975
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
         TabIndex        =   10
         Top             =   7320
         Width           =   1215
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
         TabIndex        =   9
         Top             =   7320
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   11760
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   4200
         Width           =   1815
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
         TabIndex        =   7
         Top             =   4680
         Width           =   975
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
         TabIndex        =   6
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   7560
         TabIndex        =   3
         Top             =   6360
         Width           =   3135
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   7560
         TabIndex        =   2
         Top             =   6720
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Height          =   855
         Left            =   11760
         Picture         =   "temp_frame.frx":18F8A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6600
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   11760
         TabIndex        =   4
         Top             =   6000
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   41418753
         CurrentDate     =   41382
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   11760
         TabIndex        =   5
         Top             =   5280
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   41418753
         CurrentDate     =   41382
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
         TabIndex        =   36
         Top             =   3720
         Width           =   735
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
         TabIndex        =   35
         Top             =   4560
         Width           =   1455
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
         TabIndex        =   34
         Top             =   5400
         Width           =   2175
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
         TabIndex        =   33
         Top             =   3840
         Width           =   2295
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
         TabIndex        =   32
         Top             =   3840
         Width           =   1455
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
         TabIndex        =   31
         Top             =   4680
         Width           =   1815
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
         TabIndex        =   30
         Top             =   5160
         Width           =   1335
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
         TabIndex        =   29
         Top             =   6000
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
         Left            =   6000
         TabIndex        =   28
         Top             =   7320
         Width           =   975
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
         TabIndex        =   27
         Top             =   5040
         Width           =   1815
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
         TabIndex        =   26
         Top             =   5760
         Width           =   1335
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
         TabIndex        =   25
         Top             =   6360
         Width           =   1575
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
         TabIndex        =   24
         Top             =   6720
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
         TabIndex        =   23
         Top             =   6720
         Width           =   735
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000005&
         Height          =   495
         Left            =   7560
         TabIndex        =   22
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000005&
         Height          =   495
         Left            =   9240
         TabIndex        =   21
         Top             =   3960
         Width           =   1455
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
         TabIndex        =   20
         Top             =   4080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
