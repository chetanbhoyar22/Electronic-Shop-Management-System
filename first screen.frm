VERSION 5.00
Begin VB.Form frmWelcme 
   BackColor       =   &H00FFFFFF&
   Caption         =   "frmInfo"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start "
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   23.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9480
      TabIndex        =   0
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "2017 - 2018"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   23.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   8
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "( BCCA - III YEAR )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   7
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "3. Samrin Ali"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   735
      Left            =   480
      TabIndex        =   6
      Top             =   6720
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "2. Vaibhav Kharadkar"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   735
      Left            =   480
      TabIndex        =   5
      Top             =   6000
      Width           =   6615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Chetan Bhoyar"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   5280
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Group Members :-"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   26.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   4560
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ELECTRONIC SHOWROOM     MANAGEMENT SYSTEM  "
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   30
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   2520
      TabIndex        =   2
      Top             =   1560
      Width           =   8295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   975
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   6375
   End
End
Attribute VB_Name = "frmWelcme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmLogin.Show
End Sub

