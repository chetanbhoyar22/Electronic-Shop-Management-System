VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPCategory 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Maa Vaishnavi Enterprises, Wardha. : Product Categories.............."
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2475
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc adoCategory 
      Height          =   390
      Left            =   75
      Top             =   2025
      Width           =   6510
      _ExtentX        =   11483
      _ExtentY        =   688
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=Laxmi_DSN"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "Laxmi_DSN"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CATEGORIES"
      Caption         =   "Maa Vaishnavi Enterprises, Main Road, Wardha - 442 001 MS"
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
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5370
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4050
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "No Update"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2730
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1410
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      Height          =   1200
      Left            =   90
      TabIndex        =   7
      Top             =   60
      Width           =   6540
      Begin VB.TextBox Text2 
         DataField       =   "PRODUCT_CATEGORY_DESCRIPTION"
         DataSource      =   "adoCategory"
         Height          =   330
         Left            =   2250
         MaxLength       =   50
         TabIndex        =   2
         Top             =   690
         Width           =   4200
      End
      Begin VB.TextBox Text1 
         DataField       =   "PRODUCT_CATEGORY"
         DataSource      =   "adoCategory"
         Height          =   330
         Left            =   2250
         MaxLength       =   30
         TabIndex        =   1
         Top             =   270
         Width           =   4200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Category Descrip."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   150
         TabIndex        =   9
         Top             =   750
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   315
         Width           =   2025
      End
   End
End
Attribute VB_Name = "frmPCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
adoCategory.Recordset.AddNew
Cntrls_Enable True
End Sub

Public Sub Cntrls_Enable(Ena As Boolean)
'for New : ena =True

Frame1.Enabled = Ena
Command2.Enabled = Ena
Command3.Enabled = Ena

Command1.Enabled = Not Ena
Command4.Enabled = Not Ena

adoCategory.Enabled = Not Ena

End Sub

Private Sub Command2_Click()
If Text1.Text <> "" Then
    Dim RSc As Recordset
    Set RSc = New Recordset
    RSc.Open "select * from CATEGORIES where PRODUCT_CATEGORY='" & Text1.Text & "'", adoCategory.ConnectionString, adOpenDynamic, adLockPessimistic, adCmdText
    If RSc.BOF Then
        adoCategory.Recordset.Update
        Cntrls_Enable False
    Else
        MsgBox "This Product Category already Entered........."
        Text1.SetFocus
    End If
Else
    MsgBox "Data can not be Saved, Due to Incomplete Information"
End If
End Sub

Private Sub Command3_Click()
adoCategory.Recordset.CancelUpdate
Cntrls_Enable False
End Sub

Private Sub Command4_Click()
If Not adoCategory.Recordset.BOF And Not adoCategory.Recordset.EOF Then
    If MsgBox("You want to remove this Record?", vbQuestion + vbOKCancel, "Are You Sure?") = vbOK Then
        adoCategory.Recordset.Delete
        MsgBox "Record removed successfully.............."
        adoCategory.Refresh
    End If
End If
End Sub

Private Sub Command5_Click()
Unload Me
End Sub
