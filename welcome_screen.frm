VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
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
   Picture         =   "welcome_screen.frx":3D1A
   ScaleHeight     =   311
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   749
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3615
      Left            =   5640
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "Proceed"
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
         Left            =   3960
         TabIndex        =   16
         Top             =   3120
         Width           =   1335
      End
      Begin VB.DriveListBox Drive1 
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
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   2535
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   240
         Picture         =   "welcome_screen.frx":17093
         ScaleHeight     =   435
         ScaleWidth      =   450
         TabIndex        =   12
         Top             =   120
         Width           =   450
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   2760
         TabIndex        =   11
         Top             =   600
         Width           =   2535
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "The database is missing. Please navigate to the data base."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "ERROR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   9240
      Picture         =   "welcome_screen.frx":176EB
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   8400
      Top             =   240
   End
   Begin MSAdodcLib.Adodc Database 
      Height          =   375
      Left            =   3360
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Flight reservation System\Swift Airlines.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Flight reservation System\Swift Airlines.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Flights"
      Caption         =   "Database"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8400
      Top             =   240
   End
   Begin VB.CommandButton Command3 
      Height          =   975
      Left            =   8400
      Picture         =   "welcome_screen.frx":17F81
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Administrator"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Height          =   975
      Left            =   9600
      Picture         =   "welcome_screen.frx":18376
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Operator"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Log In"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   8160
      TabIndex        =   0
      ToolTipText     =   "Choose your privilage"
      Top             =   2640
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5640
      Picture         =   "welcome_screen.frx":18881
      ScaleHeight     =   480
      ScaleWidth      =   600
      TabIndex        =   20
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "loading.."
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
      Left            =   6360
      TabIndex        =   18
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Swift Airlines Reservation Portal"
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
      Left            =   5760
      TabIndex        =   17
      Top             =   1920
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000000&
      Height          =   735
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   5415
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
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
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
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "decembere"
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
      Left            =   6600
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5760
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim usertp, rc, tck As Integer
Dim pth, cpth As String
Dim st(0 To 3) As String


Private Sub Command1_Click()

On Error GoTo Erl
Database.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pth + ";Persist Security Info=False"
Database.Refresh
MsgBox "Database selected!"
Form_Load
Exit Sub

Erl:
MsgBox "Please select a proper database!"
End Sub

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

Private Sub Command4_Click()
Me.Enabled = False
frmAbout.Show
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.List(Drive1.ListIndex)
End Sub

Private Sub File1_Click()
pth = Dir1.Path + "\" + File1.List(File1.ListIndex)
End Sub

Private Sub Form_Load()
Database.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Dir1.Path + "\" + "Swift Airlines.mdb" + ";Persist Security Info=False"

Timer2.Enabled = True
Frame2.Visible = False
On Error GoTo err

Database.Refresh

Set conn = New ADODB.Connection
    conn.Open Me.cs
Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic

tck = 0

st(0) = "Welcome to Swift Airlines Reservation Portal"

rs.Source = "SELECT * FROM Location"
rs.Open
    
    If rs.EOF <> True Then
    rs.MoveFirst
    End If
    
    rc = 0
    
    While rs.EOF <> True
        rc = rc + 1
        If rs("LOC_DIST") = 0 Then
            Label11.Caption = rs("LOC_NAME")
        End If
        rs.MoveNext
    Wend
    
rs.Close

st(1) = "Currently we are booking tickets for " + Str(rc) + " Location(s)"

st(2) = "Hastle Free easy booking and we are up 24 * 7"

st(3) = "Have a nice and safe journey"

Exit Sub

err:
Timer2.Enabled = False
  Frame2.Visible = True
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

Public Property Get cs() As String
cs = Database.ConnectionString
End Property

Private Sub Timer2_Timer()

If tck = 0 Then
tck = 1
ElseIf tck = 1 Then
tck = 2
ElseIf tck = 2 Then
tck = 3
ElseIf tck = 3 Then
tck = 0
End If

Label10.Caption = st(tck)
End Sub
