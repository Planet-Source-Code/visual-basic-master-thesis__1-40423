VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMBROWSESTOCKS 
   Caption         =   "Item(s)"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   Icon            =   "FRMBROWSESTOCKS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CMDCLOSE 
      Caption         =   "Close"
      Height          =   435
      Left            =   6600
      TabIndex        =   2
      Top             =   3240
      Width           =   1320
   End
   Begin VB.CommandButton CMDSELECT 
      Caption         =   "Select"
      Height          =   435
      Left            =   5160
      TabIndex        =   1
      Top             =   3240
      Width           =   1320
   End
   Begin MSComctlLib.ListView LSTSTOCKS 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item Number"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Unitprice"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Re_ord_level"
         Object.Width           =   2293
      EndProperty
   End
End
Attribute VB_Name = "FRMBROWSESTOCKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim List As ListItem

Private Sub CMDCANCEL_Click()
    Unload Me
End Sub

Private Sub CMDCLOSE_Click()
Unload Me
End Sub

Private Sub CMDSELECT_Click()
 Dim Qty As Long
 Dim PQuery As String

If AddStocks = "Item" Then
    
    Qty = GetQuantity()
    If Qty = 0 Then Exit Sub
    Set List = FRMREPLACEMENT.LSTITEM.FindItem(LSTSTOCKS.SelectedItem.Text, , , lvwPartial)
    If List Is Nothing Then
        Set List = FRMREPLACEMENT.LSTITEM.ListItems.Add(, , LSTSTOCKS.SelectedItem.Text, , 0)
        With List
          .SubItems(1) = LSTSTOCKS.SelectedItem.SubItems(1)
          .SubItems(2) = Qty
          .SubItems(3) = Format(Now, "mm/dd/yyyy")
         
            End With
    
    Else
        List.SubItems(2) = List.SubItems(2) + Qty
    End If
    
    
'Set List = LSTSTOCKS.FindItem(LSTSTOCKS.SelectedItem.Text, , , lvwPartial)
'With List
    '.SubItems(3) = .SubItems(3) - Qty
'End With
'UpdateStocksQuantity LSTSTOCKS.SelectedItem.SubItems(3), LSTSTOCKS.SelectedItem.Text
AddStocks = ""
FRMREPLACEMENT.CMDADD.Enabled = True
FRMREPLACEMENT.CMDREMOVE1.Enabled = True
Unload Me

Else
 
    Qty = GetQuantity()
    If Qty = 0 Then Exit Sub
    Set List = FRMREPLACEMENT.LSTREP.FindItem(LSTSTOCKS.SelectedItem.Text, , , lvwPartial)
    If List Is Nothing Then
        Set List = FRMREPLACEMENT.LSTREP.ListItems.Add(, , LSTSTOCKS.SelectedItem.Text, , 0)
        With List
          .SubItems(1) = LSTSTOCKS.SelectedItem.SubItems(1)
          .SubItems(2) = Qty
          .SubItems(3) = Format(Now, "mm/dd/yyyy")
         
            End With
    
    Else
        List.SubItems(2) = List.SubItems(2) + Qty
    End If
    
    
Set List = LSTSTOCKS.FindItem(LSTSTOCKS.SelectedItem.Text, , , lvwPartial)
With List
    .SubItems(3) = .SubItems(3) - Qty
End With
UpdateStocksQuantity LSTSTOCKS.SelectedItem.SubItems(3), LSTSTOCKS.SelectedItem.Text
AddStocks = ""
Unload Me
End If
End Sub


Private Sub Form_Load()
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Set TQR = DBMain.QueryDefs("StockDetails")
    Set TRS = TQR.OpenRecordset()
    SetFlatList LSTSTOCKS, Me
    DisplayRecord LSTSTOCKS, TRS, 4, 0
End Sub

Private Sub LSTSTOCKS_DblClick()
    CMDSELECT_Click
End Sub


