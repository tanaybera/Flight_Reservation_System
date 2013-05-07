VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "login_dialouge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   120
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conn As ADODB.Connection
Public rs As ADODB.Recordset

Private Sub cmdCancel_Click()
    Form1.Enabled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If txtUserName = "" Or txtPassword = "" Then
    MsgBox "Field Empty. Please Fill Up!"
    Exit Sub
    End If
    
    rs.Source = "SELECT * FROM login WHERE userid='" + txtUserName.Text + "' AND password = '" + txtPassword + "' AND usertype=" + Str(Form1.usertype)
    rs.Open
    If rs.RecordCount > 0 Then
        If Form1.usertype = 1 Then
        Admin.Show
        Unload Me
        ElseIf Form1.usertype = 2 Then
        Form3.Show
        Unload Me
        End If
        Form1.Hide
    Else
        MsgBox "Invalid Credentials, try again!", , "Login"
        txtPassword.SetFocus
    End If
    rs.Close
End Sub

Private Sub Form_Load()
If (Form1.usertype = 1) Then
Me.Caption = "Administrator Login"
End If
If (Form1.usertype = 2) Then
Me.Caption = "Operator Login"
End If


Set conn = New ADODB.Connection
    conn.Open Form1.cs
Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
End Sub

