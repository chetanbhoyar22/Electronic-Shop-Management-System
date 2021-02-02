VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrders 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Maa Vaishnavi Enterprises, Wardha. : Order Details................."
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Cancel Order"
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
      Left            =   1920
      TabIndex        =   10
      Top             =   4590
      Width           =   1770
   End
   Begin MSAdodcLib.Adodc adoCategory 
      Height          =   360
      Left            =   4935
      Top             =   -420
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
   Begin MSAdodcLib.Adodc adoProduct 
      Height          =   330
      Left            =   240
      Top             =   -480
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
      RecordSource    =   "SELECT PRODUCT_ID, PRODUCT_COMPANY_ID, PRODUCT_NAME, PRODUCT_RATE FROM PRODUCTS;"
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
   Begin MSAdodcLib.Adodc adoOrders 
      Height          =   390
      Left            =   120
      Top             =   5160
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
      RecordSource    =   "ORDERS"
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
      BackColor       =   &H008080FF&
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
      Left            =   5400
      TabIndex        =   11
      Top             =   4590
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Place Order"
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
      Left            =   120
      TabIndex        =   9
      Top             =   4590
      Width           =   1770
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   495
      Left            =   -9000
      TabIndex        =   12
      Top             =   4590
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   4455
      Left            =   90
      TabIndex        =   14
      Top             =   45
      Width           =   6540
      Begin VB.TextBox Text7 
         DataField       =   "ORDER_COMPANY_ID"
         DataSource      =   "adoOrders"
         Height          =   315
         Left            =   -7000
         TabIndex        =   26
         Text            =   "Text7"
         Top             =   795
         Width           =   795
      End
      Begin VB.TextBox Text4 
         DataField       =   "DD_NO"
         DataSource      =   "adoOrders"
         Height          =   330
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   8
         Top             =   4050
         Width           =   4200
      End
      Begin VB.TextBox Text3 
         DataField       =   "BANK_NAME"
         DataSource      =   "adoOrders"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2220
         MaxLength       =   30
         TabIndex        =   7
         Top             =   3630
         Width           =   4200
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CASH"
         DataField       =   "CASH_BANK"
         DataSource      =   "adoOrders"
         Height          =   420
         Left            =   2220
         TabIndex        =   6
         Top             =   3180
         Width           =   1560
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   2220
         TabIndex        =   0
         Top             =   240
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   635
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   224919555
         CurrentDate     =   38406
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmOrders.frx":0000
         DataField       =   "ORDER_PRODUCT_ID"
         DataSource      =   "adoOrders"
         Height          =   315
         Left            =   2220
         TabIndex        =   2
         Top             =   1110
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "PRODUCT_NAME"
         BoundColumn     =   "PRODUCT_ID"
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox Text6 
         DataField       =   "ORDER_QTY"
         DataSource      =   "adoOrders"
         Height          =   330
         Left            =   2220
         MaxLength       =   3
         TabIndex        =   4
         Top             =   2370
         Width           =   4200
      End
      Begin VB.TextBox Text5 
         DataField       =   "PRODUCT_RATE"
         DataSource      =   "adoOrders"
         Height          =   330
         Left            =   2220
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1950
         Width           =   4200
      End
      Begin VB.TextBox Text2 
         DataField       =   "ORDER_AMOUNT"
         DataSource      =   "adoOrders"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2220
         TabIndex        =   5
         Top             =   2790
         Width           =   4200
      End
      Begin VB.TextBox Text1 
         DataField       =   "ORDER_DATE"
         DataSource      =   "adoOrders"
         Height          =   330
         Left            =   -7000
         TabIndex        =   13
         Top             =   240
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmOrders.frx":0019
         DataField       =   "ORDER_PRODUCT_CATEGORY"
         DataSource      =   "adoOrders"
         Height          =   315
         Left            =   2220
         TabIndex        =   1
         Top             =   690
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
         Caption         =   "DD/Cheque No."
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
         Index           =   9
         Left            =   150
         TabIndex        =   25
         Top             =   4065
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Mode"
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
         Index           =   8
         Left            =   150
         TabIndex        =   24
         Top             =   3255
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
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
         Index           =   7
         Left            =   150
         TabIndex        =   23
         Top             =   3660
         Width           =   2025
      End
      Begin VB.Label Label2 
         Height          =   375
         Left            =   2220
         TabIndex        =   22
         Top             =   1560
         Width           =   3840
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Company"
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
         Index           =   6
         Left            =   150
         TabIndex        =   21
         Top             =   1560
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity Ordered"
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
         Index           =   5
         Left            =   150
         TabIndex        =   20
         Top             =   2430
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Rate"
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
         Index           =   4
         Left            =   150
         TabIndex        =   19
         Top             =   2025
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
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
         Index           =   3
         Left            =   150
         TabIndex        =   18
         Top             =   1170
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
         Index           =   2
         Left            =   150
         TabIndex        =   17
         Top             =   750
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         TabIndex        =   16
         Top             =   2835
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Order Date"
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
         TabIndex        =   15
         Top             =   300
         Width           =   2025
      End
   End
End
Attribute VB_Name = "frmOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NayaRec As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Text3.Text = ""
    Text4.Text = ""
    Text3.Enabled = False
    Text4.Enabled = False
Else
    Text3.Text = ""
    Text4.Text = ""
    Text3.Enabled = True
    Text4.Enabled = True
    If Text3.Enabled = True Then Text3.SetFocus
End If

End Sub

Private Sub Command1_Click()
adoOrders.Recordset.AddNew
Cntrls_Enable True
End Sub

Public Sub Cntrls_Enable(Ena As Boolean)
'for New : ena =True
NayaRec = Ena

Frame1.Enabled = Ena
Command2.Enabled = Ena
Command3.Enabled = Ena

Command1.Enabled = Not Ena
'Command4.Enabled = Not Ena

adoOrders.Enabled = Not Ena

End Sub

Private Sub Command2_Click()
Text1.Text = CDate(DTPicker1.Value)
If IsDate(Text1.Text) And DataCombo1.Text <> "" And DataCombo2.Text <> "" And Val(Text5.Text) > 0 And Val(Text6.Text) > 0 Then
    If Check1.Value = 1 Then
        adoOrders.Recordset.Update
        MsgBox "Order Successfully Placed.........."
        Cntrls_Enable False
        Command1_Click
    Else
        If Trim(Text3.Text) <> "" And Trim(Text4.Text) <> "" Then
            adoOrders.Recordset.Update
            MsgBox "Order Successfully Placed.........."
            Cntrls_Enable False
            Command1_Click
        Else
            MsgBox "Bank Details Not filled Properly........"
            If Text3.Enabled = True Then Text3.SetFocus
        End If
    End If
Else
    MsgBox "Data can not be Saved, Due to Incomplete Information"
End If
End Sub

Private Sub Command3_Click()
adoOrders.Recordset.CancelUpdate
Cntrls_Enable False
Command1_Click
End Sub

Private Sub Command4_Click()
If Not adoOrders.Recordset.BOF And Not adoOrders.Recordset.EOF Then
    If MsgBox("You want to remove this Record?", vbQuestion + vbOKCancel, "Are You Sure?") = vbOK Then
        adoOrders.Recordset.Delete
        MsgBox "Record removed successfully.............."
        adoOrders.Refresh
    End If
End If
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub DataCombo1_Change()
If DataCombo1.Text <> "" Then
    Dim REC As Recordset
    Set REC = New Recordset
    adoProduct.Recordset.MoveFirst
        
    adoProduct.Recordset.Find "PRODUCT_ID=" & DataCombo1.BoundText
    Text7.Text = adoProduct.Recordset.Fields(1)
    Text5.Text = Val(adoProduct.Recordset.Fields(3)) - (Val(adoProduct.Recordset.Fields(3)) * 10 / 100)
    
    REC.Open "SELECT COMPANY_NAME FROM COMPANIES WHERE COMPANY_ID=" & adoProduct.Recordset.Fields(1), adoProduct.ConnectionString, adOpenDynamic, adLockPessimistic, adCmdText
    
    Label2.Caption = REC.Fields(0)
Else
    Label2.Caption = ""
End If
End Sub

Private Sub DataCombo2_Change()
If DataCombo2.Text <> "" Then
    
    adoProduct.RecordSource = "SELECT PRODUCT_ID, PRODUCT_COMPANY_ID, PRODUCT_NAME, PRODUCT_RATE FROM PRODUCTS WHERE PRODUCT_CATEGORY='" & Trim(DataCombo2.Text) & "'"
    adoProduct.Refresh
    
    If adoProduct.Recordset.BOF Then
        DataCombo1.Enabled = False
        DataCombo2.SetFocus
    Else
        DataCombo1.Enabled = True
        DataCombo1.ReFill
        DataCombo1.Text = ""
        If DataCombo1.Enabled = True Then DataCombo1.SetFocus
    End If
End If
End Sub

Private Sub DTPicker1_Click()
If NayaRec = True Then
    If CDate(Format(DTPicker1.Value, "dd/mm/yyyy")) > CDate(Format(Now, "dd/mm/yyyy")) Then DTPicker1.Value = Now
    Text1.Text = CDate(DTPicker1.Value)
End If
End Sub

Private Sub Form_Load()
NayaRec = False
Command1_Click
End Sub

Private Sub Text5_Change()
Text2.Text = CStr(Val(Text5.Text) * Val(Text6.Text))
End Sub

Private Sub Text6_Change()
Text2.Text = CStr(Val(Text5.Text) * Val(Text6.Text))
End Sub
