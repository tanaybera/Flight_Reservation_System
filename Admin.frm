VERSION 5.00
Begin VB.Form Admin 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrative Control Panel, Swift Airlines Pvt LTD   © [ 2011 - 2015 ]"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11745
   Icon            =   "Admin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   318
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   783
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command12 
      BackColor       =   &H80000005&
      Height          =   495
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Admin.frx":3D1A
      Style           =   1  'Graphical
      TabIndex        =   101
      ToolTipText     =   "Tickets"
      Top             =   1320
      Width           =   615
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   240
      Picture         =   "Admin.frx":434C
      ScaleHeight     =   750
      ScaleWidth      =   900
      TabIndex        =   104
      Top             =   360
      Width           =   900
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H80000005&
      DownPicture     =   "Admin.frx":4CC0
      Height          =   495
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Admin.frx":5298
      Style           =   1  'Graphical
      TabIndex        =   102
      ToolTipText     =   "Logout"
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Height          =   495
      Left            =   120
      Picture         =   "Admin.frx":58B6
      Style           =   1  'Graphical
      TabIndex        =   94
      ToolTipText     =   "Home"
      Top             =   1320
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
      Begin VB.CommandButton Command3 
         Height          =   495
         Left            =   120
         Picture         =   "Admin.frx":6AAB
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000005&
         Height          =   495
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Admin.frx":A434
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Height          =   495
         Left            =   120
         Picture         =   "Admin.frx":E108
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   120
         Picture         =   "Admin.frx":11C50
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Add New Flight"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   2280
      TabIndex        =   47
      Top             =   120
      Width           =   9375
      Begin VB.TextBox Text20 
         Height          =   375
         Left            =   3720
         TabIndex        =   108
         Top             =   1080
         Width           =   3255
      End
      Begin VB.CommandButton Command11 
         Caption         =   "NEXT"
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
         Left            =   600
         TabIndex        =   99
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H80000005&
         Caption         =   "Tue"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   91
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H80000005&
         Caption         =   "Wed"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   90
         Top             =   1800
         Width           =   735
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H80000005&
         Caption         =   "Thur"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   89
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H80000005&
         Caption         =   "Fri"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4680
         TabIndex        =   88
         Top             =   1800
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H80000005&
         Caption         =   "Sat"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5280
         TabIndex        =   87
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H80000005&
         Caption         =   "Sun"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5880
         TabIndex        =   86
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000005&
         Caption         =   "Mon"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   85
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Add Flight"
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
         Left            =   5520
         TabIndex        =   59
         Top             =   3480
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   840
         TabIndex        =   58
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4200
         TabIndex        =   57
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   56
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         TabIndex        =   55
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   54
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   53
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   52
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   51
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   50
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   49
         Top             =   3720
         Width           =   615
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
         Height          =   3630
         Left            =   7560
         MultiSelect     =   1  'Simple
         TabIndex        =   48
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Airlines :"
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
         Left            =   2760
         TabIndex        =   107
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000005&
         Caption         =   "From"
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
         Left            =   240
         TabIndex        =   75
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Caption         =   "To"
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
         Left            =   3720
         TabIndex        =   74
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000005&
         Caption         =   "Running On"
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
         Left            =   600
         TabIndex        =   73
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000005&
         Caption         =   "Departure"
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
         Left            =   600
         TabIndex        =   72
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000005&
         Caption         =   "[ hh : mm ]"
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
         Left            =   2760
         TabIndex        =   71
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000005&
         Caption         =   "Average km(s) covered in 1 hr"
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
         Left            =   3960
         TabIndex        =   70
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000005&
         Caption         =   "Max Accomodation"
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
         Left            =   600
         TabIndex        =   69
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000005&
         Caption         =   "Reservation Fare"
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
         Left            =   600
         TabIndex        =   68
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000005&
         Caption         =   "First"
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
         Left            =   2400
         TabIndex        =   67
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000005&
         Caption         =   "Travel Class"
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
         Left            =   600
         TabIndex        =   66
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000005&
         Caption         =   "Business"
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
         Left            =   3360
         TabIndex        =   65
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000005&
         Caption         =   "Economy"
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
         Left            =   4320
         TabIndex        =   64
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000005&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   1920
         TabIndex        =   63
         Top             =   3000
         Width           =   135
      End
      Begin VB.Label Label21 
         BackColor       =   &H80000005&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   2040
         TabIndex        =   62
         Top             =   4200
         Width           =   135
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000005&
         Caption         =   "Reservation Fare must be on 'per unit distance' basis"
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
         Left            =   2160
         TabIndex        =   61
         Top             =   4200
         Width           =   3975
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000005&
         Caption         =   "Select Intermediate Stops"
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
         Left            =   7320
         TabIndex        =   60
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000005&
      Caption         =   "Add New Location"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   2280
      TabIndex        =   76
      Top             =   120
      Width           =   9375
      Begin VB.CommandButton Command14 
         Caption         =   "Confirm"
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
         Left            =   6840
         TabIndex        =   121
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2880
         TabIndex        =   83
         Text            =   "Select Current Destination"
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   79
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   78
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Add Location"
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
         Left            =   4080
         TabIndex        =   77
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Select Current Location"
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
         Left            =   240
         TabIndex        =   84
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Destination Name"
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
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "Distance"
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
         Left            =   5880
         TabIndex        =   81
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "The distance of the new destination must be relative to the current geographical location"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   80
         Top             =   720
         Width           =   7695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000005&
      Caption         =   "Edit Flight Protocols"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   2280
      TabIndex        =   9
      Top             =   120
      Width           =   9375
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   110
         Top             =   3360
         Width           =   2175
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
         Height          =   2160
         Left            =   7800
         MultiSelect     =   1  'Simple
         TabIndex        =   30
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6480
         TabIndex        =   29
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6480
         TabIndex        =   28
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   27
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   26
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   25
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   24
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8400
         TabIndex        =   23
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   22
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H80000005&
         Caption         =   "Mon"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   21
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H80000005&
         Caption         =   "Sun"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7200
         TabIndex        =   20
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H80000005&
         Caption         =   "Sat"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6600
         TabIndex        =   19
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H80000005&
         Caption         =   "Fri"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6000
         TabIndex        =   18
         Top             =   840
         Width           =   495
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H80000005&
         Caption         =   "Thur"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5280
         TabIndex        =   17
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox Check13 
         BackColor       =   &H80000005&
         Caption         =   "Wed"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   16
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox Check14 
         BackColor       =   &H80000005&
         Caption         =   "Tue"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Update"
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
         Left            =   5880
         TabIndex        =   14
         Top             =   3360
         Width           =   1455
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label57 
         Height          =   375
         Left            =   7800
         TabIndex        =   111
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label56 
         BackColor       =   &H80000005&
         Caption         =   "Airlines"
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
         Left            =   2160
         TabIndex        =   109
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000005&
         Caption         =   "Select Intermediate Stops"
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
         Left            =   7320
         TabIndex        =   46
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000005&
         Caption         =   "Reservation Fare must be on 'per unit distance' basis"
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
         Left            =   2760
         TabIndex        =   45
         Top             =   4080
         Width           =   3975
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000005&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   2640
         TabIndex        =   44
         Top             =   4080
         Width           =   135
      End
      Begin VB.Label Label24 
         BackColor       =   &H80000005&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   3600
         TabIndex        =   43
         Top             =   2760
         Width           =   135
      End
      Begin VB.Label Label25 
         BackColor       =   &H80000005&
         Caption         =   "Economy"
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
         Left            =   6480
         TabIndex        =   42
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label26 
         BackColor       =   &H80000005&
         Caption         =   "Business"
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
         Left            =   5280
         TabIndex        =   41
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label27 
         BackColor       =   &H80000005&
         Caption         =   "Travel Class"
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
         Left            =   2160
         TabIndex        =   40
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label28 
         BackColor       =   &H80000005&
         Caption         =   "First"
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
         Left            =   4320
         TabIndex        =   39
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label29 
         BackColor       =   &H80000005&
         Caption         =   "Reservation Fare"
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
         Left            =   2160
         TabIndex        =   38
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label30 
         BackColor       =   &H80000005&
         Caption         =   "Max Accomodation"
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
         Left            =   2160
         TabIndex        =   37
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label31 
         BackColor       =   &H80000005&
         Caption         =   "Average km(s) covered in 1 hr"
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
         TabIndex        =   36
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label32 
         BackColor       =   &H80000005&
         Caption         =   "[ hh : mm ]"
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
         Left            =   4680
         TabIndex        =   35
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label33 
         BackColor       =   &H80000005&
         Caption         =   "Departure"
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
         Left            =   2280
         TabIndex        =   34
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label34 
         BackColor       =   &H80000005&
         Caption         =   "Running On"
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
         Left            =   2160
         TabIndex        =   33
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label35 
         BackColor       =   &H80000005&
         Caption         =   "To"
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
         Left            =   4920
         TabIndex        =   32
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label36 
         BackColor       =   &H80000005&
         Caption         =   "From"
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
         Left            =   2280
         TabIndex        =   31
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000005&
      Caption         =   "Tickets"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   2280
      TabIndex        =   112
      Top             =   120
      Width           =   9375
      Begin VB.CheckBox Check15 
         BackColor       =   &H80000005&
         Caption         =   "Mark as Read"
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
         Left            =   7440
         TabIndex        =   118
         Top             =   840
         Width           =   1695
      End
      Begin VB.ListBox List5 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "Admin.frx":15BED
         Left            =   120
         List            =   "Admin.frx":15BEF
         TabIndex        =   113
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "Tickets"
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
         TabIndex        =   117
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label60 
         BackColor       =   &H80000005&
         Caption         =   "----------------------------------------------------------------"
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
         Left            =   2400
         TabIndex        =   116
         Top             =   1440
         Width           =   6495
      End
      Begin VB.Label Label59 
         BackColor       =   &H80000005&
         Caption         =   "___________________________________________"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2400
         TabIndex        =   115
         Top             =   2040
         Width           =   5775
      End
      Begin VB.Label Label58 
         BackColor       =   &H80000005&
         Caption         =   "*******************************************"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   114
         Top             =   720
         Width           =   4335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000005&
      Caption         =   "Remove exixting Flights"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   9375
      Begin VB.CommandButton Command6 
         Caption         =   "Modify"
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
         TabIndex        =   5
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Confirm Delete"
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
         Left            =   2640
         TabIndex        =   2
         Top             =   2040
         Width           =   2655
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "Admin.frx":15BF1
         Left            =   120
         List            =   "Admin.frx":15BF3
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label37 
         BackColor       =   &H80000005&
         Caption         =   "The details of selected Flight will appear below. Please verify before proceeding as this action cannot be undone."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   6255
      End
      Begin VB.Label Label39 
         BackColor       =   &H80000005&
         Caption         =   "Passenger "
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
         Left            =   2880
         TabIndex        =   6
         Top             =   1440
         Width           =   4935
      End
      Begin VB.Label Label38 
         BackColor       =   &H80000005&
         Caption         =   "Flight Id : "
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
         Left            =   2880
         TabIndex        =   7
         Top             =   960
         Width           =   4935
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000005&
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   2280
      TabIndex        =   92
      Top             =   120
      Width           =   9375
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3675
         Left            =   5640
         Picture         =   "Admin.frx":15BF5
         ScaleHeight     =   3675
         ScaleWidth      =   3150
         TabIndex        =   100
         Top             =   480
         Width           =   3150
      End
      Begin VB.Label Label41 
         BackColor       =   &H80000005&
         Caption         =   "TILL NOW"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   120
         Top             =   2520
         Width           =   5535
      End
      Begin VB.Label Label40 
         BackColor       =   &H80000005&
         Caption         =   "TILL NOW"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   119
         Top             =   1920
         Width           =   5535
      End
      Begin VB.Label Label42 
         BackColor       =   &H80000005&
         Caption         =   "TILL NOW"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   93
         Top             =   1320
         Width           =   5535
      End
   End
   Begin VB.Label Label54 
      BackStyle       =   0  'Transparent
      Caption         =   "unread"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   106
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label53 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1320
      TabIndex        =   105
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "gcjdhcsgcsudhgviushviushvis"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   960
      TabIndex        =   103
      Top             =   1440
      Width           =   375
   End
End
Attribute VB_Name = "Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ad As String
Dim adbool As Boolean
Dim cdd As Integer
Dim temp_str, cl As String
Public conn As ADODB.Connection
Public rs, es, fs, ts As ADODB.Recordset



Private Sub Check15_Click()
    
    Set conn = New ADODB.Connection
    conn.Open Form1.cs
    Set ts = New ADODB.Recordset
    ts.ActiveConnection = conn
    ts.CursorLocation = adUseClient
    ts.CursorType = adOpenDynamic
    ts.LockType = adLockOptimistic
        
    
    ts.Source = "SELECT * FROM tickets WHERE ID = " + Str(List5.ItemData(List5.ListIndex))
    ts.Open
        If Check15 = 1 Then
            ts("mbit") = True
        Else
            ts("mbit") = False
        End If
    ts.Update
    ts.Close
    
    
End Sub

Private Sub Command12_Click()

Label58.Caption = ""
Label59.Caption = ""
Label60.Caption = "Pick a Ticket to view details"
Check15.Visible = False
rs.Source = "SELECT * FROM tickets ORDER BY mbit DESC"
rs.Open
    
    If rs.EOF <> True Then
    rs.MoveFirst
    End If
    Dim rc As Integer
    rc = 0
    List5.Clear
    While rs.EOF <> True
        List5.AddItem (rs("cname"))
        List5.ItemData(rc) = rs("ID")
        rc = rc + 1
        rs.MoveNext
    Wend
    
rs.Close

Frame1.Visible = False
Frame3.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame4.Visible = False
Frame7.Visible = True


End Sub

Private Sub Command14_Click()

rs.Source = "SELECT * FROM Location"
rs.Open

If rs.EOF <> True Then
rs.MoveFirst
End If

    While rs.EOF <> True
        rs("LOC_DIST") = rs("LOC_DIST") - Combo5.ItemData(Combo5.ListIndex)
        rs.Update
        rs.MoveNext
    Wend
rs.Close
MsgBox Combo5.List(Combo5.ListIndex) + " was successfully set as current location"
cl = Combo5.List(Combo5.ListIndex)

Command1_Click

End Sub

Private Sub Form_QueryUnload(cancel As Integer, unloadmode As Integer)

MsgBox "You have been logged out successfully!"
Form1.Show
Form1.Enabled = True
Unload Me
End Sub



Private Sub Combo3_Change()
rs.Source = "SELECT * FROM location WHERE LOC_NAME = '" + Combo4.Text + "' OR LOC_NAME = '" + Combo3.Text + "'"
        
        rs.Open
        
        If rs.RecordCount <> 2 Then
            
        Else
           Dim desfrom, desto As String
           Dim disfrom, disto As Integer
           desfrom = Combo4.Text
           desto = Combo3.Text
           adbool = False
           
           
           rs.Filter = "LOC_NAME LIKE '" + desto + "'"
           disto = rs.Fields("LOC_DIST")
           rs.Filter = "LOC_NAME LIKE '" + desfrom + "'"
           disfrom = rs.Fields("LOC_DIST")
           rs.Close
           
            
                If disfrom < disto Then
                    rs.Source = "SELECT * FROM Location ORDER BY LOC_DIST ASC"
                    rs.Open
                    
                    rs.MoveFirst
                    adbool = True
                    'MsgBox "SELECT * FROM Location WHERE LOC_DIST BETWEEN '" + Str(disfrom) + "' AND '" + Str(disto) + "' ORDER BY LOC_DIST ASC==" + Str(rs.RecordCount)
                    rs.Filter = "LOC_DIST >= '" + Str(disfrom) + "' AND LOC_DIST <= " + Str(disto)
                    List1.Clear
                    While rs.EOF <> True
                    If rs.Fields("LOC_NAME") <> desfrom And rs.Fields("LOC_NAME") <> desto Then
                    List1.AddItem (rs.Fields("LOC_NAME"))
                    
                    End If
                    rs.MoveNext
                    Wend
                    
                Else
                                
                    rs.Source = "SELECT * FROM Location ORDER BY LOC_DIST DESC"
                    rs.Open
                    
                    rs.MoveFirst
                    adbool = True
                    'MsgBox "SELECT * FROM Location WHERE LOC_DIST BETWEEN '" + Str(disfrom) + "' AND '" + Str(disto) + "' ORDER BY LOC_DIST ASC==" + Str(rs.RecordCount)
                    rs.Filter = "LOC_DIST <= '" + Str(disfrom) + "' AND LOC_DIST >= " + Str(disto)
                    List1.Clear
                    While rs.EOF <> True
                    If rs.Fields("LOC_NAME") <> desfrom And rs.Fields("LOC_NAME") <> desto Then
                    List1.AddItem (rs.Fields("LOC_NAME"))
                    End If
                    rs.MoveNext
                    Wend
                End If
            
        End If
        
        rs.Close
        
End Sub

Private Sub Combo4_Change()
rs.Source = "SELECT * FROM location WHERE LOC_NAME = '" + Combo4.Text + "' OR LOC_NAME = '" + Combo3.Text + "'"
        
        rs.Open
        
        If rs.RecordCount <> 2 Then
            
        Else
           Dim desfrom, desto As String
           Dim disfrom, disto As Integer
           desfrom = Combo4.Text
           desto = Combo3.Text
           adbool = False
           
           
           rs.Filter = "LOC_NAME LIKE '" + desto + "'"
           disto = rs.Fields("LOC_DIST")
           rs.Filter = "LOC_NAME LIKE '" + desfrom + "'"
           disfrom = rs.Fields("LOC_DIST")
           rs.Close
           
            
                If disfrom < disto Then
                    rs.Source = "SELECT * FROM Location ORDER BY LOC_DIST ASC"
                    rs.Open
                    
                    rs.MoveFirst
                    adbool = True
                    'MsgBox "SELECT * FROM Location WHERE LOC_DIST BETWEEN '" + Str(disfrom) + "' AND '" + Str(disto) + "' ORDER BY LOC_DIST ASC==" + Str(rs.RecordCount)
                    rs.Filter = "LOC_DIST >= '" + Str(disfrom) + "' AND LOC_DIST <= " + Str(disto)
                    List1.Clear
                    While rs.EOF <> True
                    If rs.Fields("LOC_NAME") <> desfrom And rs.Fields("LOC_NAME") <> desto Then
                    List1.AddItem (rs.Fields("LOC_NAME"))
                    
                    End If
                    rs.MoveNext
                    Wend
                    
                Else
                                
                    rs.Source = "SELECT * FROM Location ORDER BY LOC_DIST DESC"
                    rs.Open
                    
                    rs.MoveFirst
                    adbool = True
                    'MsgBox "SELECT * FROM Location WHERE LOC_DIST BETWEEN '" + Str(disfrom) + "' AND '" + Str(disto) + "' ORDER BY LOC_DIST ASC==" + Str(rs.RecordCount)
                    rs.Filter = "LOC_DIST <= '" + Str(disfrom) + "' AND LOC_DIST >= " + Str(disto)
                    List1.Clear
                    While rs.EOF <> True
                    If rs.Fields("LOC_NAME") <> desfrom And rs.Fields("LOC_NAME") <> desto Then
                    List1.AddItem (rs.Fields("LOC_NAME"))
                    End If
                    rs.MoveNext
                    Wend
                End If
            
        End If
        
        rs.Close
        
End Sub

Private Sub Command1_Click()

    rs.Source = "SELECT * FROM Location"
    rs.Open
    If rs.EOF <> True Then
    rs.MoveFirst
    End If
    Dim rc As Integer
    rc = 0
    
    While rs.EOF <> True
    Combo5.AddItem (rs.Fields("LOC_NAME"))
    Combo5.ItemData(rc) = rs("LOC_DIST")
    rc = rc + 1
    rs.MoveNext
    Wend
    
    Combo5.Text = cl
    rs.Close
    
Frame1.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame5.Visible = True
End Sub


Private Sub Command10_Click()
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame1.Visible = False
Frame7.Visible = False
Frame6.Visible = True

rs.Source = "SELECT * FROM reservation"
rs.Open
    
    If rs.EOF <> True Then
    rs.MoveFirst
    End If
    Dim rc As Integer
    rc = 0
    
    While rs.EOF <> True
        rc = rc + 1
        rs.MoveNext
    Wend
    
rs.Close
    
Label42.Caption = " TILL NOW " + Str(rc) + " TICKET(S) HAVE BEEN SOLD!"

rs.Source = "SELECT * FROM Location"
rs.Open
    
    If rs.EOF <> True Then
    rs.MoveFirst
    End If
    
    rc = 0
    
    While rs.EOF <> True
        rc = rc + 1
        If rs("LOC_DIST") = 0 Then
            cl = rs("LOC_NAME")
        End If
        rs.MoveNext
    Wend
    
rs.Close

Label40.Caption = " We have reached " + Str(rc) + " Location(s) so far.  Currently @ " + cl


rs.Source = "SELECT * FROM flights"
rs.Open
    
    If rs.EOF <> True Then
    rs.MoveFirst
    End If
    
    rc = 0
    
    While rs.EOF <> True
        rc = rc + 1
        rs.MoveNext
    Wend
    
rs.Close
    
Label41.Caption = " " + Str(rc) + " Flight(s) are registtered and running."

End Sub

Private Sub Command11_Click()
    
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    
    rs.Source = "SELECT * FROM location WHERE LOC_NAME = '" + Combo1.Text + "' OR LOC_NAME = '" + Combo2.Text + "'"
    
    rs.Open
    rs.Requery
    
    If rs.RecordCount <> 2 Then
        MsgBox "Please Choose a valid destinations first"
        rs.Close
        Exit Sub
    Else
       Dim desfrom, desto As String
       Dim disfrom, disto As Integer
       desfrom = Combo1.Text
       desto = Combo2.Text
       adbool = False
       
       
       rs.Filter = "LOC_NAME LIKE '" + desto + "'"
       disto = rs.Fields("LOC_DIST")
       rs.Filter = "LOC_NAME LIKE '" + desfrom + "'"
       disfrom = rs.Fields("LOC_DIST")
       rs.Close
       
        
            If disfrom < disto Then
                rs.Source = "SELECT * FROM Location ORDER BY LOC_DIST ASC"
                rs.Open
                
                rs.MoveFirst
                adbool = True
                'MsgBox "SELECT * FROM Location WHERE LOC_DIST BETWEEN '" + Str(disfrom) + "' AND '" + Str(disto) + "' ORDER BY LOC_DIST ASC==" + Str(rs.RecordCount)
                rs.Filter = "LOC_DIST >= '" + Str(disfrom) + "' AND LOC_DIST <= " + Str(disto)
                List2.Clear
                While rs.EOF <> True
                If rs.Fields("LOC_NAME") <> desfrom And rs.Fields("LOC_NAME") <> desto Then
                List2.AddItem (rs.Fields("LOC_NAME"))
                
                End If
                rs.MoveNext
                Wend
                
            Else
                            
                rs.Source = "SELECT * FROM Location ORDER BY LOC_DIST DESC"
                rs.Open
                
                rs.MoveFirst
                adbool = True
                'MsgBox "SELECT * FROM Location WHERE LOC_DIST BETWEEN '" + Str(disfrom) + "' AND '" + Str(disto) + "' ORDER BY LOC_DIST ASC==" + Str(rs.RecordCount)
                rs.Filter = "LOC_DIST <= '" + Str(disfrom) + "' AND LOC_DIST >= " + Str(disto)
                List2.Clear
                While rs.EOF <> True
                If rs.Fields("LOC_NAME") <> desfrom And rs.Fields("LOC_NAME") <> desto Then
                List2.AddItem (rs.Fields("LOC_NAME"))
                End If
                rs.MoveNext
                Wend
            End If
        
    End If
    
rs.Close


If adbool Then

Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True
Check5.Enabled = True
Check6.Enabled = True
Check7.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text20.Enabled = True
List2.Enabled = True
Command8.Enabled = True

End If




End Sub

Private Sub Command13_Click()
Unload Me
End Sub

Private Sub Command2_Click()

Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
Check5.Enabled = False
Check6.Enabled = False
Check7.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text20.Enabled = False
List2.Enabled = False
Command8.Enabled = False



rs.Source = "SELECT * FROM Location"
rs.Open
    
    If rs.RecordCount > 2 Then
        rs.MoveFirst
        Combo1.Clear
        Combo2.Clear
        While rs.EOF <> True
        Combo1.AddItem (rs.Fields("LOC_NAME"))
        Combo2.AddItem (rs.Fields("LOC_NAME"))
        rs.MoveNext
        Wend
    Else
        MsgBox "Theres no or one location stored in the database. Please add destinations first"
        rs.Close
        Command1_Click
    End If
    
rs.Close

Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame1.Visible = True
End Sub

Private Sub Command3_Click()



Combo4.Enabled = False
Combo3.Enabled = False
Check8.Enabled = False
Check9.Enabled = False
Check10.Enabled = False
Check11.Enabled = False
Check12.Enabled = False
Check13.Enabled = False
Check14.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
Text17.Enabled = False
Text18.Enabled = False
Text21.Enabled = False
Command7.Enabled = False
List1.Enabled = False












Set conn = New ADODB.Connection
    conn.Open Form1.cs
    Set fs = New ADODB.Recordset
    fs.ActiveConnection = conn
    fs.CursorLocation = adUseClient
    fs.CursorType = adOpenStatic 'adOpenDynamic
    fs.LockType = adLockOptimistic
    
    fs.Source = "SELECT * FROM Flights"
    
    fs.Open
    fs.Requery
    
    List3.Clear
    
    If fs.EOF <> True Then
    fs.MoveFirst
    Else
    MsgBox "no data in your records"
    End If
    Dim lt As Integer
    lt = 0
    While fs.EOF <> True
        List3.AddItem (fs.Fields("AIRLINES") + " - " + Str(fs.Fields("FIGHT_ID")))
        List3.ItemData(lt) = fs.Fields("FIGHT_ID")
        fs.MoveNext
        lt = lt + 1
    Wend



    
rs.Source = "SELECT * FROM Location"
rs.Open
    
    If rs.RecordCount > 2 Then
        rs.MoveFirst
        Combo3.Clear
        Combo4.Clear
        While rs.EOF <> True
        Combo3.AddItem (rs.Fields("LOC_NAME"))
        Combo4.AddItem (rs.Fields("LOC_NAME"))
        rs.MoveNext
        Wend
    Else
        MsgBox "Theres no or one location stored in the database. Please add destinations first"
        rs.Close
        Command1_Click
    End If
    
rs.Close






Frame1.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame3.Visible = True
End Sub

Private Sub Command4_Click()

Label38.Caption = "*******************************"
Label39.Caption = "Please select a record from the List"
Set conn = New ADODB.Connection
    conn.Open Form1.cs
    Set fs = New ADODB.Recordset
    fs.ActiveConnection = conn
    fs.CursorLocation = adUseClient
    fs.CursorType = adOpenDynamic
    fs.LockType = adLockOptimistic
    
    fs.Source = "SELECT * FROM Flights"
    
    fs.Open
    fs.Requery
    
    List4.Clear
    
    If fs.EOF <> True Then
    fs.MoveFirst
    Else
    MsgBox "no data in your records"
    End If
    Dim lt As Integer
    lt = 0
    While fs.EOF <> True
        List4.AddItem (fs.Fields("AIRLINES") + " - " + Str(fs.Fields("FIGHT_ID")))
        List4.ItemData(lt) = fs.Fields("FIGHT_ID")
        fs.MoveNext
        lt = lt + 1
    Wend
Frame1.Visible = False
Frame3.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame4.Visible = True
End Sub

Private Sub Command5_Click()
If fs.EOF = True Then
MsgBox "Select a flight first!"
Exit Sub
End If


If MsgBox("Delete " + fs.Fields("AIRLINES") + " [ " + Str(fs.Fields("FIGHT_ID")) + " ] ?", vbExclamation + vbOKCancel, "Confirm Delete") = vbOK Then
fs.Delete
Command4_Click
Exit Sub
End If
End Sub

Private Sub Command6_Click()
Frame4.Visible = False
Frame3.Visible = True
End Sub

Private Sub Command7_Click()

Set conn = New ADODB.Connection
    conn.Open Form1.cs
Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    
    
    If Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Or Text16.Text = "" Or Text17.Text = "" Or Text18.Text = "" Or Text21.Text = "" Then
    MsgBox "Fields Empty! Please fill up!!"
    Exit Sub
    End If
    
    rs.Source = "SELECT * FROM flights WHERE FIGHT_ID = " + Label57.Caption
    rs.Open
    
    rs("AIRLINES") = Text21.Text
    
    rs("mon") = Check8
    rs("tue") = Check14
    rs("wed") = Check13
    rs("thur") = Check12
    rs("fri") = Check11
    rs("sat") = Check10
    rs("sun") = Check9
    
    rs("FROM") = Combo4.Text
    rs("TO") = Combo3.Text
    
    rs("dept") = Text18.Text
    rs("kms") = Text17.Text
    
    rs("FMA") = Text16.Text
    rs("FRA") = Text15.Text
    
    rs("BMA") = Text14.Text
    rs("BRA") = Text13.Text
    
    rs("EMA") = Text12.Text
    rs("ERA") = Text11.Text
    
    Dim num, i, j As Integer
    Dim ls As String
    
    
    num = List1.ListCount
    i = 0
    j = 0
    ls = ""
    If num > 0 Then
    While i < num
        If List1.Selected(i) = True Then
            If j = 0 Then
            ls = List1.List(i)
            j = j + 1
            Else
            ls = ls + "_" + List1.List(i)
            End If
        End If
        i = i + 1
    Wend
    End If
    
    rs("IL") = ls
    rs.Update
    
    MsgBox "DETAILS  UPDATED SUCCESSFULLY "
    
    rs.Close
    
    Command3_Click

    

End Sub

Private Sub Command8_Click()

Set conn = New ADODB.Connection
    conn.Open Form1.cs
Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    
    
    If Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
    MsgBox "Fields Empty! Please fill up!!"
    Exit Sub
    End If
    
    rs.Source = "SELECT * FROM flights"
    rs.Open
    rs.AddNew
    
    rs("AIRLINES") = Text20.Text
    
    rs("mon") = Check1
    rs("tue") = Check7
    rs("wed") = Check6
    rs("thur") = Check5
    rs("fri") = Check4
    rs("sat") = Check3
    rs("sun") = Check2
    
    rs("FROM") = Combo1.Text
    rs("TO") = Combo2.Text
    
    rs("dept") = Text3.Text
    rs("kms") = Text4.Text
    
    rs("FMA") = Text5.Text
    rs("FRA") = Text6.Text
    
    rs("BMA") = Text7.Text
    rs("BRA") = Text8.Text
    
    rs("EMA") = Text9.Text
    rs("ERA") = Text10.Text
    
    Dim num, i, j As Integer
    Dim ls As String
    
    
    num = List2.ListCount
    i = 0
    j = 0
    ls = ""
    If num > 0 Then
    While i < num
        If List2.Selected(i) = True Then
            If j = 0 Then
            ls = List2.List(i)
            j = j + 1
            Else
            ls = ls + "_" + List2.List(i)
            End If
        End If
        i = i + 1
    Wend
    End If
    
    rs("IL") = ls
    rs.Update

    rs.Close
    
    MsgBox "Records added successfully!"

    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text20.Text = ""

    Check1 = 0
    Check2 = 0
    Check3 = 0
    Check4 = 0
    Check5 = 0
    Check6 = 0
    Check7 = 0
    
    List2.Clear
    
    Command2_Click
End Sub

Private Sub Command9_Click()

If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Please enter values to insert."
Exit Sub
End If

    rs.Source = "SELECT * FROM Location WHERE LOC_NAME LIKE '%" & Text1.Text & "'"
    rs.Open

Dim rec As Integer
rec = 0

rec = rs.RecordCount

rs.Close

If rec > 0 Then
If MsgBox("Duplicate Entry!! Already Exists! Overwrite?", vbYesNo, "Add Location") = vbYes Then
'd
Else
Exit Sub
End If
End If
    
    rs.Source = "SELECT * FROM Location"
    rs.Open
    rs.AddNew
    rs("LOC_NAME") = Text1.Text
    rs("LOC_DIST") = Text2.Text
    rs.Update
    rs.Close
    MsgBox Text1.Text & " Location Added Successfully"
    Text1.Text = ""
    Text2.Text = ""
    
    Command1_Click
    
    
End Sub

Private Sub Form_Load()


Set conn = New ADODB.Connection
    conn.Open Form1.cs
Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    
    
    rs.Source = "SELECT * FROM tickets"
    rs.Open
    
    If rs.EOF <> True Then
    rs.MoveFirst
    End If
    Dim rc As Integer
    rc = 0
    
    While rs.EOF <> True
        If Not rs("mbit") Then
        rc = rc + 1
        End If
        rs.MoveNext
    Wend
    
    Label53.Caption = Str(rc)
rs.Close

    
    
Command10_Click

End Sub





Private Sub List3_Click()

    
    fs.Filter = "FIGHT_ID LIKE '" & List3.ItemData(List3.ListIndex) & "'"
    
    
    If fs("mon") Then
    Check8 = 1
    Else
    Check8 = 0
    End If
    
    If fs("tue") Then
    Check14 = 1
    Else
    Check14 = 0
    End If
    
    If fs("wed") Then
    Check13 = 1
    Else
    Check13 = 0
    End If
    
    If fs("thur") Then
    Check12 = 1
    Else
    Check12 = 0
    End If
    
    If fs("fri") Then
    Check11 = 1
    Else
    Check11 = 0
    End If
    
    If fs("sat") Then
    Check10 = 1
    Else
    Check10 = 0
    End If
    
    If fs("sun") Then
    Check9 = 1
    Else
    Check9 = 0
    End If
    
    
    Combo4.Text = fs("FROM")
    Combo3.Text = fs("TO")
    
    Text18.Text = fs("dept")
    Text17.Text = fs("kms")
    
    Text16.Text = fs("FMA")
    Text15.Text = fs("FRA")
    
    Text14.Text = fs("BMA")
    Text13.Text = fs("BRA")
    
    Text12.Text = fs("EMA")
    Text11.Text = fs("ERA")
    
    Text21.Text = fs("AIRLINES")
    Label57.Caption = fs("FIGHT_ID")
    
        rs.Source = "SELECT * FROM location WHERE LOC_NAME = '" + Combo4.Text + "' OR LOC_NAME = '" + Combo3.Text + "'"
        
        rs.Open
        
           Dim desfrom, desto As String
           Dim disfrom, disto As Integer
           desfrom = Combo4.Text
           desto = Combo3.Text
           adbool = False
           
           
           rs.Filter = "LOC_NAME LIKE '" + desto + "'"
           disto = rs.Fields("LOC_DIST")
           rs.Filter = "LOC_NAME LIKE '" + desfrom + "'"
           disfrom = rs.Fields("LOC_DIST")
           rs.Close
           
            
                If disfrom < disto Then
                    rs.Source = "SELECT * FROM Location ORDER BY LOC_DIST ASC"
                    rs.Open
                    
                    rs.MoveFirst
                    adbool = True
                    'MsgBox "SELECT * FROM Location WHERE LOC_DIST BETWEEN '" + Str(disfrom) + "' AND '" + Str(disto) + "' ORDER BY LOC_DIST ASC==" + Str(rs.RecordCount)
                    rs.Filter = "LOC_DIST >= '" + Str(disfrom) + "' AND LOC_DIST <= " + Str(disto)
                    List1.Clear
                    While rs.EOF <> True
                    If rs.Fields("LOC_NAME") <> desfrom And rs.Fields("LOC_NAME") <> desto Then
                    List1.AddItem (rs.Fields("LOC_NAME"))
                    
                    End If
                    rs.MoveNext
                    Wend
                    
                Else
                                
                    rs.Source = "SELECT * FROM Location ORDER BY LOC_DIST DESC"
                    rs.Open
                    
                    rs.MoveFirst
                    adbool = True
                    'MsgBox "SELECT * FROM Location WHERE LOC_DIST BETWEEN '" + Str(disfrom) + "' AND '" + Str(disto) + "' ORDER BY LOC_DIST ASC==" + Str(rs.RecordCount)
                    rs.Filter = "LOC_DIST <= '" + Str(disfrom) + "' AND LOC_DIST >= " + Str(disto)
                    List1.Clear
                    While rs.EOF <> True
                    If rs.Fields("LOC_NAME") <> desfrom And rs.Fields("LOC_NAME") <> desto Then
                    List1.AddItem (rs.Fields("LOC_NAME"))
                    End If
                    rs.MoveNext
                    Wend
                End If
            
                
        rs.Close
        
        Dim tmp() As String
        tmp = Split(fs.Fields("IL"), "_")
        
        Dim nn, nnn, j, i As Integer
        nn = UBound(tmp)
        nnn = List1.ListCount
        
        For i = 0 To nn
            For j = 0 To nnn
                If List1.List(j) = tmp(i) Then
                List1.Selected(j) = True
                Exit For
                End If
            Next
        Next
        
        
        
        
        
        
        
        
        
Combo4.Enabled = True
Combo3.Enabled = True
Check8.Enabled = True
Check9.Enabled = True
Check10.Enabled = True
Check11.Enabled = True
Check12.Enabled = True
Check13.Enabled = True
Check14.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text14.Enabled = True
Text15.Enabled = True
Text16.Enabled = True
Text17.Enabled = True
Text18.Enabled = True
Text21.Enabled = True
Command7.Enabled = True
List1.Enabled = True

        
        
        
        
        
        
        
        
        
        
        
End Sub

Private Sub List4_Click()

    fs.Filter = "FIGHT_ID LIKE '" & List4.ItemData(List4.ListIndex) & "'"
    
    Label38.Caption = "Flight ID :" + Str(fs.Fields("FIGHT_ID")) + ",  airlines owned by " + fs.Fields("AIRLINES")
    Label39.Caption = "Flight between " + fs.Fields("FROM") + " to " + fs.Fields("TO")

End Sub



Private Sub List5_Click()
    rs.Source = "SELECT * FROM tickets WHERE ID = " + Str(List5.ItemData(List5.ListIndex))
    rs.Open
    
        If rs.EOF <> True Then
            Label58.Caption = "Customer Name : " + rs("cname")
            Label60.Caption = "Contact details : " + rs("cmail") + ",  " + rs("cnum")
            Label59.Caption = "Message : " + rs("mess")
            If rs("mbit") Then
                Check15 = 1
            Else
                Check15 = 0
            End If
            Check15.Visible = True
        End If
    rs.Close
End Sub

Private Sub Text19_Change()
    
    Set conn = New ADODB.Connection
    conn.Open Form1.cs
    Set es = New ADODB.Recordset
    es.ActiveConnection = conn
    es.CursorLocation = adUseClient
    es.CursorType = adOpenDynamic
    es.LockType = adLockOptimistic
    
    es.Source = "SELECT * FROM Flights WHERE AIRLINES LIKE '%" & Text19.Text & "%'"
    
    es.Open
    es.Requery
    
    List3.Clear
    
    If es.EOF <> True Then
    es.MoveFirst
    Else
    MsgBox "no data in your records"
    
        Combo4.Enabled = False
        Combo3.Enabled = False
        Check8.Enabled = False
        Check9.Enabled = False
        Check10.Enabled = False
        Check11.Enabled = False
        Check12.Enabled = False
        Check13.Enabled = False
        Check14.Enabled = False
        Text11.Enabled = False
        Text12.Enabled = False
        Text13.Enabled = False
        Text14.Enabled = False
        Text15.Enabled = False
        Text16.Enabled = False
        Text17.Enabled = False
        Text18.Enabled = False
        Text21.Enabled = False
        Command7.Enabled = False
        List1.Enabled = False

    
    
    
    End If
    
    Dim lt As Integer
    lt = 0

    
    While es.EOF <> True
        List3.AddItem (es.Fields("AIRLINES") + " - " + Str(es.Fields("FIGHT_ID")))
        List3.ItemData(lt) = es.Fields("FIGHT_ID")
        es.MoveNext
        lt = lt + 1
    Wend

    es.Close
    
End Sub

Private Sub Text22_Change()
    Set conn = New ADODB.Connection
    conn.Open Form1.cs
    Set fs = New ADODB.Recordset
    fs.ActiveConnection = conn
    fs.CursorLocation = adUseClient
    fs.CursorType = adOpenDynamic
    fs.LockType = adLockOptimistic
    
    fs.Source = "SELECT * FROM Flights WHERE AIRLINES LIKE '%" & Text22.Text & "%'"
    
    fs.Open
    fs.Requery
    
    List4.Clear
    
    If fs.EOF <> True Then
    fs.MoveFirst
    Else
    MsgBox "no data in your records"
    End If
    
    Dim lt As Integer
    lt = 0

    
    While fs.EOF <> True
        List4.AddItem (fs.Fields("AIRLINES") + " - " + Str(fs.Fields("FIGHT_ID")))
        List4.ItemData(lt) = fs.Fields("FIGHT_ID")
        fs.MoveNext
        lt = lt + 1
    Wend
   
    
End Sub
