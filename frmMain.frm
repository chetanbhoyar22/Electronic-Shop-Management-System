VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000E&
   Caption         =   "Maa Vaishnavi Enterprises, Wardha. "
   ClientHeight    =   7485
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   13350
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu nuMaster 
      Caption         =   "Master Entries"
      Begin VB.Menu mnuCompDetails 
         Caption         =   "Company Details"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProdCat 
         Caption         =   "Product Categories"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProducts 
         Caption         =   "Product Deatails"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "Transactions"
      Begin VB.Menu mnuPOrder 
         Caption         =   "Place Product Order"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuROrder 
         Caption         =   "Received Order Details"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBill 
         Caption         =   "Customer Bills"
      End
   End
   Begin VB.Menu mnuRepo 
      Caption         =   "Reports"
      Begin VB.Menu mnuRepoStiock 
         Caption         =   "Stock Report"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuBill_Click()
frmBills.Show
End Sub

Private Sub mnuCompDetails_Click()
frmCompanies.Show
End Sub

Private Sub mnuPOrder_Click()
frmOrders.Show
End Sub

Private Sub mnuProdCat_Click()
frmPCategory.Show
End Sub

Private Sub mnuProducts_Click()
frmProducts.Show
End Sub

Private Sub mnuRepoStiock_Click()
DRepoStock.Show
End Sub

Private Sub mnuROrder_Click()
frmRecdOrder.Show
End Sub
