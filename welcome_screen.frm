VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Swift Airlines Pvt Ltd    © 2011 - 2015"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11235
   Icon            =   "welcome_screen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "welcome_screen.frx":164A
   ScaleHeight     =   311
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   749
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8400
      Top             =   240
   End
   Begin VB.CommandButton Command3 
      Height          =   975
      Left            =   8520
      Picture         =   "welcome_screen.frx":149C3
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Administrator"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Height          =   975
      Left            =   9720
      Picture         =   "welcome_screen.frx":14DB8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Operator"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Log In"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   8280
      TabIndex        =   0
      ToolTipText     =   "Choose your privilage"
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "WELCOME"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000005&
      Caption         =   "AM"
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
      Left            =   10320
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000005&
      Caption         =   "55"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   20.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "55 : 55 PM"
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
      Left            =   9960
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "2013"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "decembere"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   3840
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim usertp As Integer

Private Sub Command2_Click()
usertp = 2
Load frmLogin
frmLogin.Show
Me.Enabled = False
End Sub

Private Sub Command3_Click()
usertp = 1
Load frmLogin
frmLogin.Show
Me.Enabled = False
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Format(Now, "dd")
Label2.Caption = Format(Now, "mmmm")
Label3.Caption = Format(Now, "yyyy")
Label4.Caption = Format(Now, "hh : mm AM/PM")
Label5.Caption = Format(Now, "ss")
Label6.Caption = Format(Now, "AM/PM")

End Sub

Public Property Get usertype() As Integer
usertype = usertp
End Property
