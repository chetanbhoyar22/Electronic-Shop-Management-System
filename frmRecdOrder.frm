VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRecdOrder 
   BackColor       =   &H0080FF80&
   Caption         =   "Maa Vaishnavi Enterprises, Wardha. : Received Product Order Details................."
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc adoRecdOrders 
      Height          =   390
      Left            =   120
      Top             =   3930
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
      RecordSource    =   "STOCK"
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
      Left            =   5400
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000007&
      Caption         =   "Stock Recd."
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
      TabIndex        =   4
      Top             =   3360
      Width           =   1770
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   495
      Left            =   -9000
      TabIndex        =   6
      Top             =   4590
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Enabled         =   0   'False
      Height          =   3270
      Left            =   90
      TabIndex        =   8
      Top             =   45
      Width           =   6540
      Begin VB.TextBox Text3 
         DataField       =   "STOCK_PRODUCT_ID"
         DataSource      =   "adoRecdOrders"
         Height          =   315
         Left            =   -7000
         TabIndex        =   20
         Text            =   "Text3"
         Top             =   1155
         Width           =   795
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2235
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   705
         Width           =   3285
      End
      Begin VB.TextBox Text7 
         DataField       =   "STOCK_AGAINST_ORDER_NO"
         DataSource      =   "adoRecdOrders"
         Height          =   315
         Left            =   -7000
         TabIndex        =   17
         Text            =   "Text7"
         Top             =   795
         Width           =   795
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
         CustomFormat    =   "ddd dd/MMMM/yyyy"
         Format          =   223215619
         CurrentDate     =   38406
      End
      Begin VB.TextBox Text6 
         DataField       =   "STOCK_PRODUCT_QTY"
         DataSource      =   "adoRecdOrders"
         Height          =   330
         Left            =   2220
         MaxLength       =   3
         TabIndex        =   2
         Top             =   2370
         Width           =   4200
      End
      Begin VB.TextBox Text5 
         DataField       =   "RECD_PRODUCT_RATE"
         DataSource      =   "adoRecdOrders"
         Height          =   330
         Left            =   2220
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1950
         Width           =   4200
      End
      Begin VB.TextBox Text2 
         DataField       =   "RECD_PRODUCT_AMOUNT"
         DataSource      =   "adoRecdOrders"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2220
         TabIndex        =   3
         Top             =   2790
         Width           =   4200
      End
      Begin VB.TextBox Text1 
         DataField       =   "STOCK_RECD_DATE"
         DataSource      =   "adoRecdOrders"
         Height          =   330
         Left            =   -7000
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         Height          =   375
         Left            =   2205
         TabIndex        =   18
         Top             =   1080
         Width           =   3840
      End
      Begin VB.Label Label2 
         Height          =   375
         Left            =   2220
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   2445
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
         TabIndex        =   13
         Top             =   1995
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
         TabIndex        =   12
         Top             =   1170
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Order Details"
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   2850
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Received Date"
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
         TabIndex        =   9
         Top             =   300
         Width           =   2025
      End
   End
End
Attribute VB_Name = "frmRecdOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NayaRec As Boolean
Dim Rs As Recordset
Dim OQty As Long

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

Private Sub Combo1_Click()
Rs.MoveFirst
Rs.Find "ORDER_NO=" & Combo1.ItemData(Combo1.ListIndex)
Text7.Text = Combo1.ItemData(Combo1.ListIndex)

Label2.Caption = Rs.Fields(2)
Label3.Caption = Rs.Fields(4)
OQty = Rs.Fields(5)
Text6.Text = Rs.Fields(5)
Text5.Text = Rs.Fields(6)
Text3.Text = Rs.Fields(7)

End Sub

Private Sub Command1_Click()
adoRecdOrders.Recordset.AddNew
Cntrls_Enable True
End Sub

Public Sub Cntrls_Enable(Ena As Boolean)
'for New : ena =True
NayaRec = Ena

Frame1.Enabled = Ena
Command2.Enabled = Ena
'Command3.Enabled = Ena

Command1.Enabled = Not Ena
'Command4.Enabled = Not Ena

adoRecdOrders.Enabled = Not Ena

End Sub

Private Sub Command2_Click()
Text1.Text = CDate(DTPicker1.Value)
If IsDate(Text1.Text) And Val(Text5.Text) > 0 And Val(Text6.Text) > 0 Then
        adoRecdOrders.Recordset.Update
        Dim Cn As Connection
        Set Cn = New Connection
        Cn.Open "dsn=Laxmi_DSN"
        If OQty = Val(Text6.Text) Then
            Cn.Execute "UPDATE ORDERS SET ORDERS.ORDER_RECD = 1, ORDERS.ORDER_QTY = ORDERS.ORDER_QTY -" & Val(Text6.Text) & " WHERE ORDERS.ORDER_NO=" & Val(Text7.Text)
        Else
            Cn.Execute "UPDATE ORDERS SET ORDERS.ORDER_RECD = 0, ORDERS.ORDER_QTY = ORDERS.ORDER_QTY -" & Val(Text6.Text) & " WHERE ORDERS.ORDER_NO=" & Val(Text7.Text)
        End If
        
        Cn.Execute "UPDATE PRODUCTS SET PRODUCTS.PRODUCT_STOCK = PRODUCTS.PRODUCT_STOCK + " & Val(Text6.Text) & " WHERE PRODUCTS.PRODUCT_ID=" & Val(Text3.Text)
        
        MsgBox "Received Stock Entry Successfully Entered......", , "Stock Updated"
        Cntrls_Enable False
        
        Dim STr As String
        STr = "SELECT ORDERS.ORDER_NO,'Order No. : ' + cstr(ORDERS.ORDER_NO)+' - '+ COMPANIES.COMPANY_NAME as ORD, COMPANIES.COMPANY_NAME, ORDERS.ORDER_DATE, PRODUCTS.PRODUCT_NAME, ORDERS.ORDER_QTY, ORDERS.PRODUCT_RATE, ORDERS.ORDER_PRODUCT_ID FROM ORDERS, PRODUCTS,COMPANIES WHERE ORDERS.ORDER_PRODUCT_ID = PRODUCTS.PRODUCT_ID and ORDERS.ORDER_COMPANY_ID = COMPANIES.COMPANY_ID and ORDERS.ORDER_RECD=0"
        Rs.Close
        Rs.Open STr, adoRecdOrders.ConnectionString, adOpenDynamic, adLockPessimistic, adCmdText
        Combo1.Clear
        If Not Rs.BOF Then
            While Not Rs.EOF
                Combo1.AddItem Rs.Fields(1)
                Combo1.ItemData(Combo1.NewIndex) = Rs.Fields(0)
                Rs.MoveNext
            Wend
            Combo1.ListIndex = 0
        Else
            MsgBox "No Order in Pending.........."
            Unload Me
        End If
        
        
        Command1_Click
Else
    MsgBox "Data can not be Saved, Due to Incomplete Information"
End If
End Sub

Private Sub Command3_Click()
adoRecdOrders.Recordset.CancelUpdate
Cntrls_Enable False
Command1_Click
End Sub

Private Sub Command4_Click()
If Not adoRecdOrders.Recordset.BOF And Not adoRecdOrders.Recordset.EOF Then
    If MsgBox("You want to remove this Record?", vbQuestion + vbOKCancel, "Are You Sure?") = vbOK Then
        adoRecdOrders.Recordset.Delete
        MsgBox "Record removed successfully.............."
        adoRecdOrders.Refresh
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
Set Rs = New Recordset
Dim ST As String
ST = "SELECT ORDERS.ORDER_NO,'Order No. : ' + cstr(ORDERS.ORDER_NO)+' - '+ COMPANIES.COMPANY_NAME as ORD, COMPANIES.COMPANY_NAME, ORDERS.ORDER_DATE, PRODUCTS.PRODUCT_NAME, ORDERS.ORDER_QTY, ORDERS.PRODUCT_RATE, ORDERS.ORDER_PRODUCT_ID FROM ORDERS, PRODUCTS,COMPANIES WHERE ORDERS.ORDER_PRODUCT_ID = PRODUCTS.PRODUCT_ID and ORDERS.ORDER_COMPANY_ID = COMPANIES.COMPANY_ID and ORDERS.ORDER_RECD=0"
Rs.Open ST, adoRecdOrders.ConnectionString, adOpenDynamic, adLockPessimistic, adCmdText
If Not Rs.BOF Then
    While Not Rs.EOF
        Combo1.AddItem Rs.Fields(1)
        Combo1.ItemData(Combo1.NewIndex) = Rs.Fields(0)
        Rs.MoveNext
    Wend
    Combo1.ListIndex = 0
Else
    MsgBox "No Order in Pending.........."
    Unload Me
End If
End Sub

Private Sub Text5_Change()
Text2.Text = CStr(Val(Text5.Text) * Val(Text6.Text))
End Sub

Private Sub Text6_Change()
If NayaRec = True Then
    If Val(Text6.Text) > OQty Or Val(Text6.Text) <= 0 Then Text6.Text = OQty
    Text2.Text = CStr(Val(Text5.Text) * Val(Text6.Text))
End If
End Sub







