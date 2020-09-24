VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMPAYCUSTOMER 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer(s)"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList IMGLIST 
      Left            =   7275
      Top             =   1575
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
            Picture         =   "FRMPAYCUSTOMER.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CMDCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   7305
      TabIndex        =   2
      Top             =   660
      Width           =   1350
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   435
      Left            =   7305
      TabIndex        =   1
      Top             =   165
      Width           =   1350
   End
   Begin MSComctlLib.ListView LSTCUST 
      Height          =   3105
      Left            =   165
      TabIndex        =   0
      Top             =   165
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   5477
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
         Text            =   "Customer Number"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Lastname"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Firstname"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Address"
         Object.Width           =   6174
      EndProperty
   End
End
Attribute VB_Name = "FRMPAYCUSTOMER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TQR As DAO.QueryDef
Dim TRS As DAO.Recordset
Dim Query As String

Private Sub CMDCANCEL_Click()
    Unload Me
End Sub

Private Sub CMDOK_Click()
    Query = "SELECT custno,lastname,firstname,address,creditlimit FROM customers WHERE custno='" & LSTCUST.SelectedItem.Text & "';"
    Set TQR = DBMain.CreateQueryDef("", Query)
    Set TRS = TQR.OpenRecordset()
    With FRMPAYMENTS
        .TXTF(0).Text = TRS.Fields(0)
        .TXTF(1).Text = TRS.Fields(1)
        .TXTF(2).Text = TRS.Fields(2)
        .TXTF(3).Text = TRS.Fields(3)
        .TXTF(4).Text = TRS.Fields(4)
    End With
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    SetFlatList LSTCUST, Me
    Query = "SELECT custno,lastname,firstname,address FROM customers"
    Set TQR = DBMain.CreateQueryDef("", Query)
    Set TRS = TQR.OpenRecordset()
    DisplayRecord LSTCUST, TRS, 3, 1
End Sub


Private Sub LSTCUST_DblClick()
    CMDOK_Click
End Sub

Private Sub LSTCUST_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.Caption = "Customer(s) -" & Item.SubItems(1)
End Sub
