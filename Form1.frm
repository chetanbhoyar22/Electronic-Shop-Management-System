VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmProducts 
   Caption         =   "Laxmi Electronics, Wardha. : Products Detail..................."
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4215
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc adoCategory 
      Height          =   360
      Left            =   4935
      Top             =   -1200
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Enabled         =   0
      Connect         =   "DSN=Laxmi_DSN"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "Laxmi_DSN"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT PRODUCT_CATEGORY FROM CATEGORIES;"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc adoCompany 
      Height          =   330
      Left            =   240
      Top             =   -1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Enabled         =   0
      Connect         =   "DSN=Laxmi_DSN"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "Laxmi_DSN"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT COMPANY_ID, COMPANY_NAME FROM COMPANIES;"
      Caption         =   "Adodc1"
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
   Begin MSAdodcLib.Adodc adoProducts 
      Height          =   390
      Left            =   120
      Top             =   3795
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
      RecordSource    =   "PRODUCTS"
      Caption         =   "Laxmi Electronics, Main Road, Wardha - 442 001 MS"
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
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   3225
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   3225
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "No Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   3225
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   3225
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3225
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   3090
      Left            =   90
      TabIndex        =   11
      Top             =   60
      Width           =   6540
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Form1.frx":0000
         DataField       =   "PRODUCT_COMPANY_ID"
         DataSource      =   "adoProducts"
         Height          =   315
         Left            =   2250
         TabIndex        =   4
         Top             =   1560
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "COMPANY_NAME"
         BoundColumn     =   "COMPANY_ID"
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox Text6 
         DataField       =   "PRODUCT_STOCK"
         DataSource      =   "adoProducts"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2250
         TabIndex        =   6
         Top             =   2475
         Width           =   4200
      End
      Begin VB.TextBox Text5 
         DataField       =   "PRODUCT_RATE"
         DataSource      =   "adoProducts"
         Height          =   330
         Left            =   2250
         TabIndex        =   5
         Top             =   2010
         Width           =   4200
      End
      Begin VB.TextBox Text2 
         DataField       =   "PRODUCT_DESCRIPTION"
         DataSource      =   "adoProducts"
         Height          =   330
         Left            =   2250
         TabIndex        =   2
         Top             =   690
         Width           =   4200
      End
      Begin VB.TextBox Text1 
         DataField       =   "PRODUCT_NAME"
         DataSource      =   "adoProducts"
         Height          =   330
         Left            =   2250
         TabIndex        =   1
         Top             =   270
         Width           =   4200
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Form1.frx":0019
         DataField       =   "PRODUCT_CATEGORY"
         DataSource      =   "adoProducts"
         Height          =   315
         Left            =   2250
         TabIndex        =   3
         Top             =   1140
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "PRODUCT_CATEGORY"
         BoundColumn     =   "PRODUCT_CATEGORY"
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product in Stock"
         Height          =   300
         Index           =   5
         Left            =   150
         TabIndex        =   17
         Top             =   2550
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product MRP"
         Height          =   300
         Index           =   4
         Left            =   150
         TabIndex        =   16
         Top             =   2085
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         Height          =   300
         Index           =   3
         Left            =   150
         TabIndex        =   15
         Top             =   1590
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Category"
         Height          =   300
         Index           =   2
         Left            =   150
         TabIndex        =   14
         Top             =   1155
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Description"
         Height          =   300
         Index           =   1
         Left            =   150
         TabIndex        =   13
         Top             =   750
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         Height          =   300
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   315
         Width           =   2025
      End
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
adoProducts.Recordset.AddNew
Cntrls_Enable True
End Sub

Public Sub Cntrls_Enable(Ena As Boolean)
'for New : ena =True

Frame1.Enabled = Ena
Command2.Enabled = Ena
Command3.Enabled = Ena

Command1.Enabled = Not Ena
Command4.Enabled = Not Ena

adoProducts.Enabled = Not Ena

End Sub

Private Sub Command2_Click()
If Text1.Text <> "" And DataCombo1.Text <> "" And DataCombo2.Text <> "" And Val(Text5.Text) > 0 Then
    adoProducts.Recordset.Update
    Cntrls_Enable False
Else
    MsgBox "Data can not be Saved, Due to Incomplete Information"
End If
End Sub

Private Sub Command3_Click()
adoProducts.Recordset.CancelUpdate
Cntrls_Enable False
End Sub

Private Sub Command4_Click()
If Not adoProducts.Recordset.BOF And Not adoProducts.Recordset.EOF Then
    If MsgBox("You want to remove this Record?", vbQuestion + vbOKCancel, "Are You Sure?") = vbOK Then
        adoProducts.Recordset.Delete
        MsgBox "Record removed successfully.............."
        adoProducts.Refresh
    End If
End If
End Sub

Private Sub Command5_Click()
Unload Me
End Sub
