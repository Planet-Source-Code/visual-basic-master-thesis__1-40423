VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMOCREDIT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "Order Items in Credit"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "FRMOCREDIT.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FRACUSTOMERS 
      BorderStyle     =   0  'None
      Height          =   6165
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   11160
      Begin VB.Frame FRAORDERDETAILS 
         BackColor       =   &H80000004&
         Caption         =   "[ Order Details ]"
         ForeColor       =   &H00000000&
         Height          =   5775
         Left            =   0
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   11145
         Begin MSComctlLib.ListView LSTTSTOCKS 
            Height          =   2835
            Left            =   135
            TabIndex        =   19
            Top             =   2805
            Width           =   10875
            _ExtentX        =   19182
            _ExtentY        =   5001
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "IMGLIST"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Stock Number"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Quantity"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Amount"
               Object.Width           =   3528
            EndProperty
         End
         Begin MSComctlLib.ListView LSTTORDERS 
            Height          =   2295
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   4048
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "IMGLIST"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Order Number"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Amount Ordered"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Date Ordered"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Running Balance"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Due Date"
               Object.Width           =   3528
            EndProperty
         End
      End
      Begin MSComctlLib.ImageList IMGLIST 
         Left            =   4815
         Top             =   4560
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   17
         ImageHeight     =   17
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRMOCREDIT.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FRMOCREDIT.frx":0464
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CMDCLOSE 
         Caption         =   "Close"
         Height          =   420
         Left            =   9840
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton CMDORDERS 
         Caption         =   "Orders"
         Enabled         =   0   'False
         Height          =   420
         Left            =   9840
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame FRACUST 
         Caption         =   "[ Customer ]"
         ForeColor       =   &H00FF0000&
         Height          =   3375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9645
         Begin VB.Frame Frame1 
            Caption         =   "[ Accounts ]"
            Height          =   1515
            Left            =   135
            TabIndex        =   13
            Top             =   1755
            Width           =   9375
            Begin VB.TextBox TXTDUEAMOUNT 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   16
               Text            =   "0.00"
               Top             =   240
               Width           =   2500
            End
            Begin VB.TextBox TXTACCBAL 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   345
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   14
               Text            =   "0.00"
               Top             =   870
               Width           =   2500
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Amount Due :"
               Height          =   195
               Left            =   240
               TabIndex        =   17
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Running Balance :"
               Height          =   195
               Left            =   240
               TabIndex        =   15
               Top             =   960
               Width           =   1320
            End
         End
         Begin VB.CommandButton CMDBROWSE 
            Caption         =   "..."
            Height          =   270
            Left            =   3870
            MouseIcon       =   "FRMOCREDIT.frx":08BC
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   300
            Width           =   300
         End
         Begin VB.TextBox TXTF 
            Enabled         =   0   'False
            Height          =   360
            Index           =   2
            Left            =   1800
            TabIndex        =   11
            Top             =   1275
            Width           =   3300
         End
         Begin VB.TextBox TXTF 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   1800
            TabIndex        =   10
            Top             =   780
            Width           =   3300
         End
         Begin VB.TextBox TXTF 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   330
            Index           =   0
            Left            =   1800
            TabIndex        =   9
            Top             =   285
            Width           =   1995
         End
         Begin VB.Label LBLMESS 
            AutoSize        =   -1  'True
            Caption         =   "Firstname :"
            Height          =   195
            Index           =   2
            Left            =   250
            TabIndex        =   8
            Top             =   1365
            Width           =   765
         End
         Begin VB.Label LBLMESS 
            AutoSize        =   -1  'True
            Caption         =   "Lastname :"
            Height          =   195
            Index           =   1
            Left            =   250
            TabIndex        =   7
            Top             =   870
            Width           =   780
         End
         Begin VB.Label LBLMESS 
            AutoSize        =   -1  'True
            Caption         =   "Customer Number :"
            Height          =   195
            Index           =   0
            Left            =   250
            TabIndex        =   6
            Top             =   375
            Width           =   1350
         End
      End
      Begin MSComctlLib.ListView LSTCUSTORD 
         Height          =   2295
         Left            =   120
         TabIndex        =   2
         Top             =   3600
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "IMGLIST"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Order Number"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Amount Ordered"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Date Ordered"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Running Balance"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Due Date"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip TABS 
      Height          =   6555
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11562
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "          Customers         "
            Object.Tag             =   "Customer"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "          Orders Details         "
            Object.Tag             =   "Order"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FRMOCREDIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TRS As DAO.Recordset
Dim TQR As DAO.QueryDef
Dim STRS As DAO.Recordset
Dim Query As String

Private Sub CMDBROWSE_Click()
    FRMBROWSE.Show vbModal
End Sub

Private Sub CMDCLOSE_Click()
    Unload Me
 End Sub


Private Sub CMDORDERS_Click()
    If ReachCredit(TXTF(0).Text) Then
        MsgBox "Customer reach his/her Credit limit", vbExclamation, "Account"
        Exit Sub
    End If
    If EAccounts(TXTF(0).Text) <> 0 Then
        MsgBox "Customer has an existing account to Pay", vbExclamation, "Account"
        Exit Sub
    End If
    Dim Tmp As String
    Tmp = AutoOrder
    With FRMADDORDERS
        .Caption = "[ ORDER FORM ] Customer Number : " & TXTF(0).Text
        .Tag = TXTF(0).Text
        .LBLORDDATE = " Order Date : " & Format(Now, "mm/dd/yyyy")
        .LBLORDNO = " Order Number : " & Tmp
        .LBLORDNO.Tag = Tmp
        '.LBLDUEDATE = " Due Date : " & Format(Now + 30, "mm/dd/yyyy")
    End With
    FRMADDORDERS.Show vbModal
End Sub

Private Sub Form_Activate()
    DrawBorder 100, Me
End Sub

Private Sub Form_Load()
    SetFlatList LSTCUSTORD, Me
    SetFlatList LSTTORDERS, Me
    SetFlatList LSTTSTOCKS, Me
End Sub


Private Sub LSTTORDERS_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Query = "SELECT Stockno,quantity,amount FROM orders WHERE ordno ='" & Item.Text & "';"
    Query = "SELECT stocks.Stockno,Stocks.Description,orders.quantity,orders.amount FROM orders INNER JOIN stocks ON stocks.stockno = orders.stockno WHERE orders.ordno='" & Item.Text & "';"
    Set TQR = DBMain.CreateQueryDef("", Query)
    Set TRS = TQR.OpenRecordset()
    LSTTSTOCKS.ListItems.Clear
    DisplayRecord LSTTSTOCKS, TRS, 3, 1
End Sub

Private Sub TABS_Click()
    If TABS.SelectedItem.Tag = "Customer" Then
        FRACUST.Visible = True
        FRAORDERDETAILS.Visible = False
    Else
        FRACUST.Visible = False
        FRAORDERDETAILS.Visible = True
        LSTTORDERS.ListItems.Clear
        Query = "SELECT ordno,amount,[date],balance,[due_date] FROM accounts WHERE custno='" & TXTF(0).Text & "';"
        Set TQR = DBMain.CreateQueryDef("", Query)
        Set TRS = TQR.OpenRecordset()
        DisplayRecord LSTTORDERS, TRS, 4, 1
    End If
End Sub

Private Sub TXTF_Change(Index As Integer)
    If Index = 0 Then
        If TXTF(0).Text = "" Then
            CMDORDERS.Enabled = False
            Exit Sub
        Else
            CMDORDERS.Enabled = True
        End If
        
        Change TXTF(0).Text
     DueAmount TXTF(0).Text
     End If
End Sub


