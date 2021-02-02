VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmBills 
   BackColor       =   &H0080C0FF&
   Caption         =   "Maa Vaishnavi Enterprises, Main Road, Wardha - 442 001 MS"
   ClientHeight    =   5880
   ClientLeft      =   4440
   ClientTop       =   2400
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   9915
   Begin MSAdodcLib.Adodc adoBills 
      Height          =   390
      Left            =   1155
      Top             =   6090
      Width           =   3240
      _ExtentX        =   5715
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
      RecordSource    =   "BILLS"
      Caption         =   ""
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
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   360
      Left            =   1905
      TabIndex        =   26
      Top             =   1950
      Width           =   3705
      Begin VB.TextBox txtDiscount 
         DataField       =   "BILL_DISCOUNT"
         DataSource      =   "adoBills"
         Height          =   330
         Left            =   0
         TabIndex        =   27
         Top             =   15
         Width           =   3615
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1140
      Left            =   1920
      TabIndex        =   23
      Top             =   1620
      Width           =   3690
      Begin VB.TextBox txtNetBill 
         DataField       =   "BILL_NET_AMOUNT"
         DataSource      =   "adoBills"
         Height          =   330
         Left            =   0
         TabIndex        =   25
         Top             =   735
         Width           =   3630
      End
      Begin VB.TextBox txtBillAmt 
         DataField       =   "BILL_AMOUNT"
         DataSource      =   "adoBills"
         Height          =   330
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command4 
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
      Left            =   7920
      TabIndex        =   22
      Top             =   5265
      Width           =   1920
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Final the Bill"
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
      Left            =   7320
      TabIndex        =   29
      Top             =   4695
      Width           =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Discount"
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
      Left            =   5730
      TabIndex        =   12
      Top             =   4695
      Width           =   1560
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   2880
      Left            =   45
      TabIndex        =   10
      Top             =   2880
      Width           =   5640
      Begin VB.CommandButton Command3 
         Caption         =   "Add to Bill List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3735
         TabIndex        =   9
         Top             =   1995
         Width           =   1815
      End
      Begin VB.TextBox Text11 
         Height          =   330
         Left            =   -1290
         TabIndex        =   28
         Top             =   675
         Width           =   900
      End
      Begin VB.TextBox Text10 
         Height          =   330
         Left            =   1890
         TabIndex        =   11
         Top             =   2430
         Width           =   1770
      End
      Begin VB.TextBox Text9 
         Height          =   330
         Left            =   1890
         TabIndex        =   8
         Top             =   1995
         Width           =   1770
      End
      Begin VB.TextBox Text8 
         Height          =   330
         Left            =   1890
         TabIndex        =   7
         Top             =   1560
         Width           =   1770
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmBills.frx":0000
         DataField       =   "ORDER_PRODUCT_ID"
         DataSource      =   "adoOrders"
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "PRODUCT_NAME"
         BoundColumn     =   "PRODUCT_ID"
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmBills.frx":0019
         DataField       =   "ORDER_PRODUCT_CATEGORY"
         DataSource      =   "adoOrders"
         Height          =   315
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "PRODUCT_CATEGORY"
         BoundColumn     =   "PRODUCT_CATEGORY"
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   41
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblStock 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   40
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblCompany 
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
         Left            =   1890
         TabIndex        =   36
         Top             =   1155
         Width           =   3450
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
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
         Left            =   120
         TabIndex        =   35
         Top             =   1155
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
         Index           =   12
         Left            =   150
         TabIndex        =   34
         Top             =   300
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
         Index           =   11
         Left            =   150
         TabIndex        =   33
         Top             =   705
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Rate"
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
         Index           =   10
         Left            =   150
         TabIndex        =   32
         Top             =   1620
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         TabIndex        =   31
         Top             =   2055
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Amount"
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
         TabIndex        =   30
         Top             =   2460
         Width           =   2025
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Height          =   2880
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   5640
      Begin VB.TextBox Text4 
         DataField       =   "CUST_TEL"
         DataSource      =   "adoBills"
         Height          =   330
         Left            =   1920
         TabIndex        =   4
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         DataField       =   "CUST_CITY"
         DataSource      =   "adoBills"
         Height          =   330
         Left            =   1920
         TabIndex        =   3
         Top             =   960
         Width           =   3630
      End
      Begin VB.TextBox Text2 
         DataField       =   "CUST_ADDRESS"
         DataSource      =   "adoBills"
         Height          =   330
         Left            =   1920
         TabIndex        =   2
         Top             =   600
         Width           =   3630
      End
      Begin VB.TextBox Text1 
         DataField       =   "CUST_NAME"
         DataSource      =   "adoBills"
         Height          =   330
         Left            =   1920
         TabIndex        =   1
         Top             =   225
         Width           =   3630
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Bill Amount"
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
         Left            =   165
         TabIndex        =   21
         Top             =   2445
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         Left            =   165
         TabIndex        =   20
         Top             =   255
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   165
         TabIndex        =   19
         Top             =   600
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
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
         Left            =   165
         TabIndex        =   18
         Top             =   945
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
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
         Left            =   165
         TabIndex        =   17
         Top             =   1350
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Amount"
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
         Left            =   165
         TabIndex        =   16
         Top             =   1710
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Given"
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
         Left            =   165
         TabIndex        =   15
         Top             =   2085
         Width           =   2025
      End
   End
   Begin MSAdodcLib.Adodc adoCategory 
      Height          =   360
      Left            =   -3200
      Top             =   3255
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
      Left            =   -3200
      Top             =   3690
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
   Begin VB.TextBox txtBillNo 
      DataField       =   "BILL_NO"
      DataSource      =   "adoBills"
      Height          =   375
      Left            =   4080
      TabIndex        =   39
      Text            =   "Text12"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblQty 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   3975
      Left            =   9015
      TabIndex        =   38
      Top             =   510
      Width           =   840
   End
   Begin VB.Label lblProduct 
      BackStyle       =   0  'Transparent
      Height          =   3930
      Left            =   5835
      TabIndex        =   37
      Top             =   495
      Width           =   3060
   End
   Begin VB.Shape Shape1 
      Height          =   4440
      Left            =   5730
      Top             =   135
      Width           =   4170
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Billed Product List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5775
      TabIndex        =   14
      Top             =   195
      Width           =   3180
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9000
      TabIndex        =   13
      Top             =   195
      Width           =   840
   End
End
Attribute VB_Name = "frmBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PID(1 To 10) As Long
Dim PRate(1 To 10) As Long
Dim PQty(1 To 10) As Currency
Dim PRDs As Integer
Dim iX As Integer

Private Sub Command1_Click()
'Label1.Caption = Label1.Caption & "Abc" & vbCrLf
Frame4.Enabled = True
txtDiscount.SetFocus
txtDiscount.SelStart = 0
txtDiscount.SelLength = Len(txtDiscount.Text)
End Sub

Private Sub Command2_Click()
doSave
If MsgBox("Do you want to print Bill?", vbYesNo + vbQuestion, "Bill Printing") = vbYes Then
    dRepoBill.Show
    Unload Me
End If
End Sub

Private Sub Command3_Click()
If Trim(Text1.Text) <> "" And Trim(Text3.Text) <> "" Then
    If Val(Text11.Text) > 0 And Val(Text8.Text) > 0 And Val(Text9.Text) > 0 Then
        If CheckPrd(CLng(Text11.Text), PRDs) = False Then
            If Val(Text9.Text) <= Val(lblStock.Caption) Then
                PRDs = PRDs + 1
                PID(PRDs) = Val(Text11.Text)
                PRate(PRDs) = Val(Text8.Text)
                PQty(PRDs) = Val(Text9.Text)
                
                lblProduct.Caption = lblProduct.Caption & DataCombo2.Text & vbCrLf
                lblQty.Caption = lblQty.Caption & Text9.Text & vbCrLf
                
                txtBillAmt.Text = Val(txtBillAmt.Text) + (PRate(PRDs) * PQty(PRDs))
                Text10.Text = (PRate(PRDs) * PQty(PRDs))
                If (Val(txtBillAmt.Text) - Val(txtDiscount.Text)) > 0 Then
                    txtNetBill.Text = Val(txtBillAmt.Text) - Val(txtDiscount.Text)
                Else
                    txtNetBill.Text = "0"
                End If
            Else
                MsgBox "No Sufficient Quantity in Hand for this Product......", vbCritical
                Text9.SetFocus
                Text9.Text = lblStock.Caption & ""
                Text9.SelStart = 0
                Text9.SelLength = Len(Text9.Text)
            End If
        Else
            MsgBox "Product already in the Bill List"
        End If
    Else
        MsgBox "Product Information is Incomplete................"
    End If
Else
    MsgBox "Enter first Customer Name & City for Billing......."
    Text1.SetFocus
End If
End Sub

Private Sub Command5_Click()
dRepoBill.Show
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub DataCombo1_Change()
If DataCombo1.Text <> "" Then
    
    adoProduct.RecordSource = "SELECT PRODUCT_ID, PRODUCT_COMPANY_ID, PRODUCT_NAME, PRODUCT_RATE FROM PRODUCTS WHERE PRODUCT_CATEGORY='" & Trim(DataCombo1.Text) & "'"
    adoProduct.Refresh
    
    If adoProduct.Recordset.BOF Then
        DataCombo2.Enabled = False
        DataCombo1.SetFocus
    Else
        DataCombo2.Enabled = True
        DataCombo2.ReFill
        DataCombo2.Text = ""
        If DataCombo2.Enabled = True Then DataCombo2.SetFocus
    End If
End If
End Sub

Private Sub DataCombo2_Click(Area As Integer)
If DataCombo2.Text <> "" Then
    Dim REC As Recordset
    Dim PRs As Recordset
    Set REC = New Recordset
    Set PRs = New Recordset
    adoProduct.Recordset.MoveFirst
        
    adoProduct.Recordset.Find "PRODUCT_ID=" & DataCombo2.BoundText
    Text11.Text = DataCombo2.BoundText
    
    
    Text8.Text = Val(adoProduct.Recordset.Fields(3))
    
    REC.Open "SELECT COMPANY_NAME FROM COMPANIES WHERE COMPANY_ID=" & adoProduct.Recordset.Fields(1), adoProduct.ConnectionString, adOpenDynamic, adLockPessimistic, adCmdText
    PRs.Open "SELECT PRODUCT_STOCK FROM PRODUCTS WHERE PRODUCT_ID=" & DataCombo2.BoundText, adoProduct.ConnectionString, adOpenDynamic, adLockPessimistic, adCmdText
    
    lblCompany.Caption = REC.Fields(0)
    lblStock.Caption = PRs.Fields(0) & ""
Else
    lblCompany.Caption = ""
End If
End Sub

Private Sub Form_Load()
prd = 0
For iX = 1 To 10
    PID(iX) = 0
    PQty(iX) = 0
    PRate(iX) = 0
Next iX

adoBills.Recordset.AddNew
txtBillAmt.Text = "0"
txtDiscount.Text = "0"
txtNetBill.Text = "0"
End Sub


Public Function CheckPrd(p As Long, tmp As Integer)
    Dim aa As Integer
    Dim Chk As Boolean
    Chk = False
    For aa = 1 To tmp
        If PID(aa) = p Then Chk = True
    Next aa
    CheckPrd = Chk
End Function



Private Sub doSave()
Dim Sx As Byte
Dim Cn As ADODB.Connection
Dim SRs As ADODB.Recordset
Dim BRs As ADODB.Recordset

Set Cn = New ADODB.Connection
Cn.ConnectionString = "DSN=Laxmi_DSN"
Cn.Open

adoBills.Recordset.Update

Set BRs = New ADODB.Recordset
Set SRs = New ADODB.Recordset
SRs.Open "SELECT * from BILL_DETAILS where 1=2", Cn, adOpenDynamic, adLockPessimistic, adCmdText
BRs.Open "SELECT MAX(BILL_NO) FROM BILLS", Cn, adOpenDynamic, adLockPessimistic, adCmdText
For Sx = 1 To PRDs
    SRs.AddNew
            SRs("BILL_NO") = BRs(0) & ""
            SRs("PRODUCT_ID") = PID(Sx) & ""
            SRs("PRODUCT_RATE") = PRate(Sx) & ""
            SRs("PRODUCT_QTY") = PQty(Sx) & ""
            SRs("PRODUCT_AMOUNT") = PRate(Sx) * PQty(Sx)
            
            Cn.Execute "UPDATE PRODUCTS SET PRODUCTS.PRODUCT_STOCK = PRODUCTS.PRODUCT_STOCK - " & PQty(Sx) & " WHERE PRODUCTS.PRODUCT_ID=" & PID(Sx)
    SRs.Update
Next Sx


End Sub

Private Sub Text8_Change()
Text10.Text = Val(Text8.Text) * Val(Text9.Text)
End Sub

Private Sub Text9_Change()
Text10.Text = Val(Text8.Text) * Val(Text9.Text)
End Sub

Private Sub txtDiscount_Change()
    If (Val(txtBillAmt.Text) - Val(txtDiscount.Text)) > 0 Then
        txtNetBill.Text = Val(txtBillAmt.Text) - Val(txtDiscount.Text)
    Else
        txtNetBill.Text = "0"
    End If
End Sub
