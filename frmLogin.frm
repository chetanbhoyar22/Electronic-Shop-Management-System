VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFC0FF&
   Caption         =   "Log In Here......."
   ClientHeight    =   2190
   ClientLeft      =   8415
   ClientTop       =   4935
   ClientWidth     =   4545
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4545
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Please Login Here"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   195
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1470
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Login"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1470
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1380
         MaxLength       =   25
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   825
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1380
         TabIndex        =   3
         Top             =   435
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   270
         TabIndex        =   2
         Top             =   885
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   1
         Top             =   495
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim RsChk As Recordset
Dim Cn As Connection

Set Cn = New Connection
Set RsChk = New Recordset
Cn.Open "dsn=Laxmi_DSN"
RsChk.Open "select * from users where User_id='" & Text1.Text & "' and password='" & Text2.Text & "'", Cn, adOpenDynamic, adLockPessimistic, adCmdText
ShowApp = False
If RsChk.BOF Then
    MsgBox "User / Password Invalid", vbOKOnly, "Authntication Error"
    frmLogin.Text2.Text = ""
    frmLogin.Text2.SetFocus
Else
    frmMain.Show
    Unload Me
End If

End Sub

Private Sub Command2_Click()
'StartExam = False
Unload frmLogin
End Sub

'
'Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Index = 1 Then
'    Label3(1).ForeColor = vbRed
'    Label3(1).FontBold = True
'Else
'Label3(1).ForeColor = vbYellow
'Label3(1).FontBold = False
'End If
'End Sub
Private Sub Form_Load()

End Sub
