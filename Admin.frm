VERSION 5.00
Begin VB.Form Admin 
   BackColor       =   &H80000005&
   Caption         =   "Administrative Control Panel, Swift Airlines Pvt LTD   © [ 2011 - 2015 ]"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11325
   Icon            =   "Admin.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   240
      Picture         =   "Admin.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   240
      Picture         =   "Admin.frx":55E7
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000005&
      Height          =   495
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Admin.frx":912F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
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
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton Command3 
         Height          =   495
         Left            =   120
         Picture         =   "Admin.frx":CE03
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   1815
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
      Height          =   2775
      Left            =   2280
      TabIndex        =   5
      Top             =   120
      Width           =   8895
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   360
         Width           =   1815
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
         Height          =   1740
         Left            =   120
         TabIndex        =   7
         Top             =   720
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
         TabIndex        =   6
         Top             =   2040
         Width           =   2655
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
         TabIndex        =   12
         Top             =   360
         Width           =   6255
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
         TabIndex        =   11
         Top             =   960
         Width           =   4935
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
         TabIndex        =   10
         Top             =   1440
         Width           =   4935
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
      Height          =   2775
      Left            =   2280
      TabIndex        =   51
      Top             =   120
      Width           =   8895
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
         Left            =   7320
         TabIndex        =   70
         Top             =   240
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
         TabIndex        =   69
         Top             =   240
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
         Left            =   4080
         TabIndex        =   68
         Top             =   240
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
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
         Left            =   240
         TabIndex        =   67
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
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
         Left            =   240
         TabIndex        =   65
         Top             =   1680
         Width           =   855
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
         Left            =   240
         TabIndex        =   64
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox Check5 
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
         Left            =   240
         TabIndex        =   63
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox Check6 
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
         Left            =   240
         TabIndex        =   62
         Top             =   2400
         Width           =   975
      End
      Begin VB.CheckBox Check7 
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
         Left            =   240
         TabIndex        =   61
         Top             =   960
         Width           =   975
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
         Left            =   3120
         TabIndex        =   60
         Top             =   840
         Width           =   1335
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
         Left            =   8040
         TabIndex        =   59
         Top             =   840
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
         Left            =   3720
         TabIndex        =   58
         Top             =   1680
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
         Left            =   3720
         TabIndex        =   57
         Top             =   2040
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
         Left            =   4680
         TabIndex        =   56
         Top             =   1680
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
         Left            =   4680
         TabIndex        =   55
         Top             =   2040
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
         Left            =   5640
         TabIndex        =   54
         Top             =   1680
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
         Left            =   5640
         TabIndex        =   53
         Top             =   2040
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
         Height          =   900
         Left            =   6600
         MultiSelect     =   1  'Simple
         TabIndex        =   52
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
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
         Left            =   240
         TabIndex        =   66
         Top             =   1440
         Width           =   1215
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
         TabIndex        =   86
         Top             =   240
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
         TabIndex        =   85
         Top             =   240
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
         Left            =   240
         TabIndex        =   84
         Top             =   600
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
         Left            =   2040
         TabIndex        =   83
         Top             =   840
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
         Left            =   4560
         TabIndex        =   82
         Top             =   840
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
         Left            =   5520
         TabIndex        =   81
         Top             =   840
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
         Left            =   2040
         TabIndex        =   80
         Top             =   1680
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
         Left            =   2040
         TabIndex        =   79
         Top             =   2040
         Width           =   1695
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
         Left            =   3720
         TabIndex        =   78
         Top             =   1320
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
         Left            =   2040
         TabIndex        =   77
         Top             =   1320
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
         Left            =   4680
         TabIndex        =   76
         Top             =   1320
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
         Left            =   5640
         TabIndex        =   75
         Top             =   1320
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
         Left            =   3480
         TabIndex        =   74
         Top             =   2040
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
         Left            =   2160
         TabIndex        =   73
         Top             =   2400
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
         Left            =   2280
         TabIndex        =   72
         Top             =   2400
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
         Left            =   6720
         TabIndex        =   71
         Top             =   2400
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
      Height          =   2775
      Left            =   2280
      TabIndex        =   87
      Top             =   120
      Width           =   8895
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
         Left            =   2400
         TabIndex        =   90
         Top             =   1560
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
         TabIndex        =   89
         Top             =   1560
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
         Left            =   6960
         TabIndex        =   88
         Top             =   2160
         Width           =   1455
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
         Left            =   480
         TabIndex        =   93
         Top             =   1560
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
         TabIndex        =   92
         Top             =   1560
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
         TabIndex        =   91
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
      Height          =   2775
      Left            =   2280
      TabIndex        =   13
      Top             =   120
      Width           =   8895
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
         Height          =   900
         Left            =   6600
         MultiSelect     =   1  'Simple
         TabIndex        =   34
         Top             =   1440
         Width           =   2175
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
         Left            =   5640
         TabIndex        =   33
         Top             =   2040
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
         Left            =   5640
         TabIndex        =   32
         Top             =   1680
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
         Left            =   4680
         TabIndex        =   31
         Top             =   2040
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
         Left            =   4680
         TabIndex        =   30
         Top             =   1680
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
         Left            =   3720
         TabIndex        =   29
         Top             =   2040
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
         Left            =   3720
         TabIndex        =   28
         Top             =   1680
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
         Left            =   8040
         TabIndex        =   27
         Top             =   960
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
         Left            =   3120
         TabIndex        =   26
         Top             =   960
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
         Left            =   2760
         TabIndex        =   25
         Top             =   720
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
         Left            =   6720
         TabIndex        =   24
         Top             =   720
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
         Left            =   6120
         TabIndex        =   23
         Top             =   720
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
         Left            =   5520
         TabIndex        =   22
         Top             =   720
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
         Left            =   4800
         TabIndex        =   21
         Top             =   720
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
         Left            =   4080
         TabIndex        =   20
         Top             =   720
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
         Left            =   3480
         TabIndex        =   19
         Top             =   720
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
         Left            =   7320
         TabIndex        =   18
         Top             =   240
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
         Height          =   1740
         Left            =   120
         TabIndex        =   17
         Top             =   720
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   240
         Width           =   1935
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
         Left            =   6720
         TabIndex        =   50
         Top             =   2400
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
         Left            =   2280
         TabIndex        =   49
         Top             =   2400
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
         Left            =   2160
         TabIndex        =   48
         Top             =   2400
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
         Left            =   3480
         TabIndex        =   47
         Top             =   2040
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
         Left            =   5640
         TabIndex        =   46
         Top             =   1320
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
         Left            =   4680
         TabIndex        =   45
         Top             =   1320
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
         Left            =   2040
         TabIndex        =   44
         Top             =   1320
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
         Left            =   3720
         TabIndex        =   43
         Top             =   1320
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
         Left            =   2040
         TabIndex        =   42
         Top             =   2040
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
         Left            =   2040
         TabIndex        =   41
         Top             =   1680
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
         Left            =   5520
         TabIndex        =   40
         Top             =   960
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
         Left            =   4560
         TabIndex        =   39
         Top             =   960
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
         Left            =   2040
         TabIndex        =   38
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label34 
         BackColor       =   &H80000005&
         Caption         =   "On"
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
         TabIndex        =   37
         Top             =   600
         Width           =   375
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
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frame1.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = True
End Sub

Private Sub Command2_Click()
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame1.Visible = True
End Sub

Private Sub Command3_Click()
Frame1.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame3.Visible = True
End Sub

Private Sub Command4_Click()
Frame1.Visible = False
Frame3.Visible = False
Frame5.Visible = False
Frame4.Visible = True
End Sub

Private Sub Command6_Click()
Frame4.Visible = False
Frame3.Visible = True
End Sub

Private Sub Form_Load()
Frame1.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = True
End Sub

