VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMDSTOCKS 
   Caption         =   "Item(s)"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMDSTOCKS.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CMDCANCEL 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   5865
      TabIndex        =   2
      Top             =   3240
      Width           =   1320
   End
   Begin VB.CommandButton CMDSELECT 
      Caption         =   "Select"
      Height          =   435
      Left            =   4440
      TabIndex        =   1
      Top             =   3240
      Width           =   1320
   End
   Begin MSComctlLib.ListView LSTSTOCKS 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item Number"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
   End
End
Attribute VB_Name = "FRMDSTOCKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDCANCEL_Click()
    Unload Me
End Sub

Private Sub CMDSELECT_Click()
    Dim List As ListItem
    Dim Qty As Long
    Qty = GetQuantity()
    If Qty = 0 Then Exit Sub
    Set List = FRMDELIVERY.LSTDEL.FindItem(LSTSTOCKS.SelectedItem.Text, , , lvwPartial)
    If List Is Nothing Then
    With LSTSTOCKS
        Set List = FRMDELIVERY.LSTDEL.ListItems.Add(, , .SelectedItem.Text, , 1)
        List.SubItems(1) = .SelectedItem.SubItems(1)
        List.SubItems(2) = Qty
    End With
    Else
        List.SubItems(2) = Qty
    End If
    Set List = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Query As String
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
   
    Query = "SELECT stockno,description FROM stocks;"
    Set TQR = DBMain.CreateQueryDef("", Query)
    Set TRS = TQR.OpenRecordset()
    DisplayRecord LSTSTOCKS, TRS, 1, 1
End Sub

Private Sub LSTSTOCKS_DblClick()
    CMDSELECT_Click
End Sub

