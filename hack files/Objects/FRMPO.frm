VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMPO 
   BackColor       =   &H8000000A&
   Caption         =   "Purchase Orders"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "FRMPO.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FRAPOPARENT 
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   315
      TabIndex        =   1
      Top             =   645
      Width           =   11370
      Begin VB.Frame FRADETAILS 
         Caption         =   "[ Purchase Order Details ]"
         ForeColor       =   &H00FF0000&
         Height          =   5940
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   11235
         Begin MSComctlLib.ImageList IMGLIST2 
            Left            =   5445
            Top             =   1050
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
                  Picture         =   "FRMPO.frx":000C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FRMPO.frx":0460
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView LSTDDETAILS 
            Height          =   3210
            Left            =   105
            TabIndex        =   22
            Top             =   2640
            Width           =   11040
            _ExtentX        =   19473
            _ExtentY        =   5662
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "IMGLIST2"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Item Number"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   12259
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Quantity Ordered"
               Object.Width           =   3528
            EndProperty
         End
         Begin MSComctlLib.ListView LSTDPO 
            Height          =   2250
            Left            =   105
            TabIndex        =   21
            Top             =   360
            Width           =   11040
            _ExtentX        =   19473
            _ExtentY        =   3969
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "IMGLIST2"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Purchase Number "
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Date Purchased"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Supplier Number"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Company Name"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Supplier Telephone #"
               Object.Width           =   3528
            EndProperty
         End
      End
      Begin VB.Frame FRASUPP 
         Caption         =   "[ Supplier Information ]"
         ForeColor       =   &H00FF0000&
         Height          =   1890
         Left            =   150
         TabIndex        =   8
         Top             =   180
         Width           =   11130
         Begin VB.TextBox TXTF 
            BackColor       =   &H00C0FFC0&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2000
            TabIndex        =   14
            Top             =   315
            Width           =   2175
         End
         Begin VB.TextBox TXTF 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2000
            TabIndex        =   13
            Top             =   825
            Width           =   4000
         End
         Begin VB.TextBox TXTF 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2000
            TabIndex        =   12
            Top             =   1320
            Width           =   4000
         End
         Begin VB.CommandButton CMDBROWSE 
            Caption         =   "..."
            Height          =   285
            Left            =   4215
            MouseIcon       =   "FRMPO.frx":08B4
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   315
            Width           =   285
         End
         Begin VB.TextBox TXTF 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   8500
            TabIndex        =   10
            Top             =   315
            Width           =   2000
         End
         Begin VB.TextBox TXTF 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   8500
            TabIndex        =   9
            Top             =   825
            Width           =   2000
         End
         Begin VB.Label LBLMES 
            AutoSize        =   -1  'True
            Caption         =   "Supplier Number :"
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   19
            Top             =   420
            Width           =   1260
         End
         Begin VB.Label LBLMES 
            AutoSize        =   -1  'True
            Caption         =   "Company :"
            Height          =   195
            Index           =   1
            Left            =   255
            TabIndex        =   18
            Top             =   945
            Width           =   750
         End
         Begin VB.Label LBLMES 
            AutoSize        =   -1  'True
            Caption         =   "Telephone # :"
            Height          =   195
            Index           =   2
            Left            =   255
            TabIndex        =   17
            Top             =   1410
            Width           =   1005
         End
         Begin VB.Label LBLMES 
            AutoSize        =   -1  'True
            Caption         =   "Date :"
            Height          =   195
            Index           =   3
            Left            =   7005
            TabIndex        =   16
            Top             =   435
            Width           =   435
         End
         Begin VB.Label LBLMES 
            AutoSize        =   -1  'True
            Caption         =   "Purchase Number :"
            Height          =   195
            Index           =   4
            Left            =   7005
            TabIndex        =   15
            Top             =   945
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "[ Purchase Order(s) ]"
         ForeColor       =   &H00FF0000&
         Height          =   3975
         Left            =   150
         TabIndex        =   2
         Top             =   2145
         Width           =   11130
         Begin VB.CommandButton CMDADD 
            Caption         =   "&Add"
            Enabled         =   0   'False
            Height          =   435
            Left            =   9840
            TabIndex        =   6
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton CMDREMOVE 
            Caption         =   "&Remove"
            Enabled         =   0   'False
            Height          =   435
            Left            =   9840
            TabIndex        =   5
            Top             =   825
            Width           =   1215
         End
         Begin VB.CommandButton CMDCANCEL 
            Caption         =   "Cancel"
            Height          =   435
            Left            =   9840
            TabIndex        =   4
            Top             =   3420
            Width           =   1215
         End
         Begin VB.CommandButton CMDOK 
            Caption         =   "OK"
            Enabled         =   0   'False
            Height          =   435
            Left            =   9840
            TabIndex        =   3
            Top             =   2940
            Width           =   1215
         End
         Begin MSComctlLib.ImageList IMGLIST 
            Left            =   6015
            Top             =   1680
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   17
            ImageHeight     =   17
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FRMPO.frx":0BBE
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView LSTPO 
            Height          =   3525
            Left            =   105
            TabIndex        =   7
            Top             =   330
            Width           =   9675
            _ExtentX        =   17066
            _ExtentY        =   6218
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Item Number "
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description "
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Quantity Ordered"
               Object.Width           =   3704
            EndProperty
         End
      End
   End
   Begin MSComctlLib.TabStrip TABS 
      Height          =   6735
      Left            =   270
      TabIndex        =   0
      Top             =   300
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   11880
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "     Purchase Order     "
            Object.Tag             =   "PO"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "          Details         "
            Object.Tag             =   "Details"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FRMPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TQR As DAO.QueryDef
Dim TRS As DAO.Recordset
Dim Query As String

Private Sub CMDADD_Click()
    FRMBSTOCKS.Show vbModal
End Sub

Private Sub CMDBROWSE_Click()
    FRMPOBROW.Show vbModal
End Sub

Private Sub CMDCANCEL_Click()
    Unload Me
End Sub

Private Sub CMDOK_Click()
    Dim x As Long
    If LSTPO.ListItems.Count = 0 Then
        MsgBox "There are no items to be Added to your Purchase Order", vbExclamation, "Purchase Order"
        Exit Sub
    End If
    If MsgBox("Are you sure you want to Order the Following Items", vbExclamation + vbYesNo, "Purchase Order") = vbYes Then
        With LSTPO
            For x = 1 To LSTPO.ListItems.Count
                PurchaseOrder .ListItems.Item(x).Text, TXTF(0).Text, .ListItems.Item(x).SubItems(2), TXTF(3).Text, TXTF(4).Text, False
            Next x
        End With
        MsgBox "All items has been successfully added to your Purchase Order", vbInformation, "Purchase Order"
    End If
    Unload Me
End Sub

Private Sub CMDREMOVE_Click()
    If LSTPO.ListItems.Count = 0 Then
        MsgBox "There is no Ordered stocks to be Remove", vbExclamation, "Remove item"
        Exit Sub
    End If
    LSTPO.ListItems.Remove LSTPO.SelectedItem.Index
End Sub


Private Sub Form_Load()
    SetFlatList LSTPO, Me
    SetFlatList LSTDPO, Me
    SetFlatList LSTDDETAILS, Me
    TXTF(3).Text = Format(Now, "mm/dd/yyyy")
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub LSTDPO_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Query = "SELECT pos.stockno,stocks.description,pos.quantity FROM pos INNER JOIN stocks ON pos.stockno = stocks.stockno WHERE pos.ponum ='" & Item.Text & "';"
    Set TQR = DBMain.CreateQueryDef("", Query)
    Set TRS = TQR.OpenRecordset()
    DisplayRecord LSTDDETAILS, TRS, 2, 1
End Sub

Private Sub TABS_Click()
    If TABS.SelectedItem.Tag = "PO" Then
        FRADETAILS.Visible = False
    Else
        FRADETAILS.Visible = True
        Query = "SELECT DISTINCT pos.ponum,pos.date,suppliers.suppno,suppliers.company,suppliers.telno FROM pos INNER JOIN suppliers ON pos.suppno = suppliers.suppno ORDER BY pos.ponum ASC;"
        Set TQR = DBMain.CreateQueryDef("", Query)
        Set TRS = TQR.OpenRecordset()
        DisplayRecord LSTDPO, TRS, 4, 2
    End If
End Sub

Private Sub TXTF_Change(Index As Integer)
    If Index = 0 Then
        If TXTF(0).Text = "" Then
            CMDADD.Enabled = False
            CMDREMOVE.Enabled = False
            CMDOK.Enabled = False
        Else
            CMDADD.Enabled = True
            CMDREMOVE.Enabled = True
            CMDOK.Enabled = True
            TXTF(4).Text = AutoPO
        End If
    End If
End Sub

 
