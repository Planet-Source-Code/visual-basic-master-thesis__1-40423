VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMLDAMAGED 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Damage Stocks"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "[ Damaged Stock(s) ]"
      ForeColor       =   &H00FF0000&
      Height          =   3855
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8895
      Begin VB.CommandButton Command1 
         Caption         =   "&Close"
         Height          =   495
         Left            =   1680
         TabIndex        =   4
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton CMDREPLACE 
         Caption         =   "&Return to Supplier"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   3120
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FRMLDAMAGED.frx":0000
         Left            =   4920
         List            =   "FRMLDAMAGED.frx":0002
         TabIndex        =   1
         Top             =   3240
         Width           =   2775
      End
      Begin MSComctlLib.ListView LSTDAMAGED 
         Height          =   2655
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Order No#"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Number"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   7408
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Select Order Number:"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   3240
         Width           =   1815
      End
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   8
      Height          =   4095
      Left            =   240
      Top             =   240
      Width           =   9135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   6
      Height          =   4335
      Left            =   120
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "FRMLDAMAGED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDREPLACE_Click()
Dim t As Long
Dim l, Query As String

If LSTDAMAGED.ListItems.Count = 0 Then
MsgBox "Record is Empty", _
vbOKOnly + vbInformation, "Return"
Exit Sub
End If

l = "UPDATE damaged SET Replaced= True WHERE stockno='" & LSTDAMAGED.SelectedItem.SubItems(1) & "';"
DBMain.Execute l


Query = "UPDATE stocks SET Quantity = Quantity + " & LSTDAMAGED.SelectedItem.SubItems(3) & " WHERE description ='" & LSTDAMAGED.SelectedItem.SubItems(2) & "';"
DBMain.Execute Query

MsgBox "Stocks Records was sucessfully updated", vbOKOnly + vbInformation, "Stocks Updated"
LSTDAMAGED.ListItems.Clear
RemoveListViewSelectedItems LSTDAMAGED
DamagedStocks
List
End Sub


Private Sub Combo1_Click()
List
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
listdamaged
End Sub
Private Sub listdamaged()
Dim TQR As DAO.QueryDef
Dim TRS As DAO.Recordset
Dim Query As String
Query = "SELECT distinct damaged.ord_no FROM damaged WHERE damaged.replaced =False;"

Set TQR = DBMain.CreateQueryDef("", Query)
Set TRS = TQR.OpenRecordset()

Do While Not TRS.EOF
Combo1.AddItem TRS.Fields("ord_no")

TRS.MoveNext
Loop

Set TQR = Nothing
Set TRS = Nothing
End Sub

Private Sub List()
Dim TQR As DAO.QueryDef
Dim TRS As DAO.Recordset
Dim Query As String
Query = "SELECT damaged.ord_no,damaged.stockno,damaged.desc,damaged.qty FROM damaged WHERE damaged.replaced = False"

Set TQR = DBMain.CreateQueryDef("", Query)
Set TRS = TQR.OpenRecordset()

DisplayRecord LSTDAMAGED, TRS, 3, 0
Set TQR = Nothing
Set TRS = Nothing
End Sub

Public Sub RemoveListViewSelectedItems(LSTMEMBERS As ListView)
Dim plngCounter As Long
For plngCounter = 1 To LSTMEMBERS.ListItems.Count
If LSTMEMBERS.ListItems.Item(plngCounter).Selected Then
LSTMEMBERS.ListItems.Remove (plngCounter)
Exit For
End If
Next
End Sub

