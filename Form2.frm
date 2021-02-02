VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3855
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   390
      Left            =   105
      Top             =   3450
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
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "No Update"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
Cntrls_Enable True
End Sub

Public Sub Cntrls_Enable(Ena As Boolean)
'for New : ena =True

Frame1.Enabled = Ena
Command2.Enabled = Ena
Command3.Enabled = Ena

Command1.Enabled = Not Ena
Command4.Enabled = Not Ena

Adodc1.Enabled = Not Ena

End Sub

Private Sub Command2_Click()
If text1.Text <> "" Then
    Adodc1.Recordset.Update
    Cntrls_Enable False
Else
    MsgBox "Data can not be Saved, Due to Incomplete Information"
End If
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.CancelUpdate
Cntrls_Enable False
End Sub

Private Sub Command4_Click()
If Not Adodc1.Recordset.BOF And Adodc1.Recordset.EOF Then
    If MsgBox("You want to remove this Record?", vbQuestion + vbOKCancel, "Are You Sure?") = vbOK Then
        Adodc1.Recordset.Delete
        MsgBox "Record removed successfully.............."
        Adodc1.Refresh
    End If
End If
End Sub

Private Sub Command5_Click()
Unload Me
End Sub
