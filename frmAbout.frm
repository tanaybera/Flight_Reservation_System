VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1200
      Picture         =   "frmAbout.frx":0342
      ScaleHeight     =   855
      ScaleWidth      =   3495
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000005&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Prakriti Poddar,        [ 111440110034 ]"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   3120
      Width           =   3525
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarique Shamsi,    [ 111440110058 ]"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   2760
      Width           =   3525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rakesh Kumar Ray,  [ 111440110040 ]"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   2400
      Width           =   3525
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanay Kumar Bera,   [ 111440110057 ]"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   2040
      Width           =   3525
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Visual Basic Project, Software Tools  CSE 4th Sem, B.TECH, NITMAS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   3045
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(cancel As Integer, unloadmode As Integer)
Form1.Enabled = True
Unload Me
End Sub
