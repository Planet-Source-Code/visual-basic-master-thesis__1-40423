VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMAIN 
   AutoShowChildren=   0   'False
   BackColor       =   &H80000016&
   Caption         =   "Prime Asia And Jewelry Shop Sales and Inventory System"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   Icon            =   "MDIMAIN.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6135
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   "12/19/02"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   2469
            MinWidth        =   2469
            TextSave        =   "2:29 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList IMGTOOL 
      Left            =   11160
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      WindowList      =   -1  'True
      Begin VB.Menu mnuItem 
         Caption         =   "&Item"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnusuppliers 
         Caption         =   "&Suppliers"
      End
      Begin VB.Menu mnufiledash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnushutdown 
         Caption         =   "&Shutdown"
      End
   End
   Begin VB.Menu mnutransactions 
      Caption         =   "&Transactions"
      Begin VB.Menu mnuOCash 
         Caption         =   "Order "
      End
      Begin VB.Menu mnudeliveries 
         Caption         =   "&Deliveries"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Reports"
      Begin VB.Menu mnuDailyreport 
         Caption         =   "Daily Sales Report"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnudash01 
         Caption         =   "-"
      End
      Begin VB.Menu mnusales 
         Caption         =   "Sales Report"
      End
      Begin VB.Menu mnuinventory 
         Caption         =   "Inventory Report"
      End
   End
   Begin VB.Menu mnuwindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuIcons 
         Caption         =   "&ArrangeIcons"
         Index           =   1
      End
      Begin VB.Menu mnutile 
         Caption         =   "&Tile Window"
      End
   End
End
Attribute VB_Name = "MDIMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*********************************************************************

Option Explicit

Private Sub CMDCUSTOMERS_Click()
 FRMCUSTOMERS.Show
End Sub

Private Sub CMDDELIVERIES_Click()
FRMDELIVERY.Show
End Sub

Private Sub CMDORDERCASH_Click()
FRMOCASH.Show
End Sub

Private Sub CMDORDERCREDIT_Click()
FRMOCREDIT.Show
End Sub

Private Sub CMDPAYMENT_Click()
FRMPAYMENTS.Show
End Sub

Private Sub CMDPO_Click()
FRMPO.Show
End Sub

Private Sub CMDSTOCKS_Click()
FRMSTOCKS.Show
End Sub

Private Sub CMDSUPPLIERS_Click()
FRMSUPPLIERS.Show
End Sub

Private Sub mnuAgent_Click(Index As Integer)
MDIForm_Load
End Sub

Private Sub Form_unload(Cancel As Integer)
IconMenu1.DeActivateForm
End Sub

Private Sub mnuAbout_Click(Index As Integer)
FRMINFO.Show
End Sub

Private Sub mnuAppendD_Click()
FRMDAMAGED.Show
FRMORD_NO.Show vbModal
End Sub

Private Sub mnucascade_Click()
    MDIMAIN.Arrange 0
End Sub

Private Sub mnucustomers_Click()
    FRMCUSTOMERS.Show
End Sub

Private Sub mnudelinquent_Click()
    With CRWReport
        Screen.MousePointer = vbHourglass
        .SelectionFormula = "{Accounts.due_date} < PrintDate And {Accounts.Balance} <> 0"
        .DataFiles(0) = App.Path + "\Database\MasterDB.mdb"
        .ReportFileName = App.Path + "\Reports\Delinquent.rpt"
        .WindowBorderStyle = crptFixedSingle
        .WindowState = crptMaximized
        .WindowTitle = "Delinquent Customer Report"
        .Action = 1
        Screen.MousePointer = vbDefault
    End With
End Sub

Private Sub mnuDailyreport_Click()
FRMDAILYR.Show
End Sub

Private Sub mnudeliveries_Click()
    FRMDELIVERY.Show
End Sub

Private Sub mnudelivery_Click()
    With CRWReport
        Screen.MousePointer = vbHourglass
        .DataFiles(0) = App.Path + "\Database\MasterDB.mdb"
        .ReportFileName = App.Path + "\Reports\Delivery.rpt"
        .WindowBorderStyle = crptFixedSingle
        .WindowState = crptMaximized
        .WindowTitle = "Deliveries Report"
        .Action = 1
        Screen.MousePointer = vbDefault
    End With

End Sub

Private Sub mnuHide_Click(Index As Integer)
MDIMAIN.WindowState = vbMinimized
End Sub

Private Sub mnuDlist_Click()
FRMLDAMAGED.Show vbModal
End Sub

Private Sub mnuIcons_Click(Index As Integer)
MDIMAIN.Arrange vbArrangeIcons
End Sub

Private Sub mnuImage_Click(Index As Integer)
FRMIMAGEVIEWER.Show
End Sub

Private Sub mnuinventory_Click()
    With CRWReport
        Screen.MousePointer = vbHourglass
        .DataFiles(0) = App.Path + "\Database\MasterDB.mdb"
        .ReportFileName = App.Path + "\Reports\Inventory.rpt"
        .WindowBorderStyle = crptFixedSingle
        .WindowState = crptMaximized
        .WindowTitle = "Inventory Report"
        .Action = 1
        Screen.MousePointer = vbDefault
    End With
End Sub

Private Sub mnuList_Click()
FRMREPLACEMENT.Show
FRMORD_NO.Show vbModal
End Sub

Private Sub mnuOCash_Click()
    FRMOCASH.Show
End Sub

Private Sub mnuOcredit_Click()
    FRMOCREDIT.Show
End Sub

Private Sub mnuoptions_Click()
    FRMOPTIONS.Show vbModal
End Sub

Private Sub mnupayments_Click()
    FRMPAYMENTS.Show vbModal
End Sub

Private Sub mnuPayroll_Click(Index As Integer)
FRMPAYROLL.Show
End Sub

Private Sub mnuPo_Click()
    With CRWReport
        Screen.MousePointer = vbHourglass
        .DataFiles(0) = App.Path + "\Database\MasterDB.mdb"
        .ReportFileName = App.Path + "\Reports\POS.RPT"
        .WindowBorderStyle = crptFixedSingle
        .WindowState = crptMaximized
        .WindowTitle = "Purchase Order Report"
        .Action = 1
        Screen.MousePointer = vbArrow
    End With
End Sub

Private Sub mnupurchase_Click()
    FRMPO.Show
End Sub

Private Sub mnuRepCustomer_Click()
    With CRWReport
        Screen.MousePointer = vbHourglass
       .DataFiles(0) = App.Path + "\Database\MasterDB.mdb"
        .ReportFileName = App.Path + "\Reports\customers.rpt"
        .WindowBorderStyle = crptFixedSingle
        .WindowState = crptMaximized
        .WindowTitle = "Customer MasterList Report"
        .Action = 1
        Screen.MousePointer = vbDefault
    End With
End Sub

Private Sub mnuRepStocks_Click()
    With CRWReport
        Screen.MousePointer = vbHourglass
        .WindowState = crptMaximized
        .WindowBorderStyle = crptFixedSingle
        .WindowTitle = "Stocks MasterList Report"
        .DataFiles(0) = App.Path + "\Database\MasterDB.mdb"
        .ReportFileName = App.Path + "\Reports\Stocks.rpt"
        .Action = 1
        Screen.MousePointer = vbDefault
    End With
End Sub

Private Sub mnuRepSuppliers_Click()
    With CRWReport
        Screen.MousePointer = vbHourglass
        .WindowState = crptMaximized
        .WindowBorderStyle = crptFixedSingle
        .WindowTitle = "Suppliers MasterList Report"
        .DataFiles(0) = App.Path + "\Database\MasterDB.mdb"
        .ReportFileName = App.Path + "\Reports\Suppliers.rpt"
        .Action = 1
        Screen.MousePointer = vbDefault
    End With
End Sub

Private Sub mnusales_Click()
    With CRWReport
        Screen.MousePointer = vbHourglass
        .WindowState = crptMaximized
        .WindowBorderStyle = crptFixedSingle
        .WindowTitle = "Sales Report"
        .DataFiles(0) = App.Path + "\Database\MasterDB.mdb"
        .ReportFileName = App.Path + "\Reports\Sales.rpt"
        .Action = 1
        Screen.MousePointer = vbDefault
    End With
End Sub

Private Sub mnushutdown_Click()
    Shutdown
End Sub

Private Sub mnustatement_Click()
    FRMPRINTCUSTOMER.Show vbModal
End Sub

Private Sub mnuItem_Click()
    FRMSTOCKS.Show
End Sub

Private Sub mnusuppliers_Click()
    FRMSUPPLIERS.Show vbModal
End Sub

Private Sub mnusupplier_Click()

End Sub

Private Sub mnutile_Click()
    MDIMAIN.Arrange 1
End Sub






