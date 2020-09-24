VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMREPLACEMENT 
   Caption         =   "Replacement"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "FRMREPLACEMENT.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   9840
      ScaleHeight     =   5715
      ScaleWidth      =   1755
      TabIndex        =   8
      Top             =   480
      Width           =   1815
      Begin VB.CommandButton CMDCLOSE 
         Caption         =   "&Close"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton CMDSAVE 
         Caption         =   "&Save"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[Replacement Item]"
      Height          =   2895
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Width           =   9495
      Begin VB.CommandButton CMDREMOVE1 
         Caption         =   "&Remove"
         Height          =   495
         Left            =   8040
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton CMDADD 
         Caption         =   "&Add"
         Height          =   495
         Left            =   8040
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin MSComctlLib.ListView LSTREP 
         Height          =   2535
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12640511
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7232
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Item for replacement]"
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9495
      Begin VB.CommandButton CMDREMOVE 
         Caption         =   "&Remove"
         Height          =   495
         Left            =   8040
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton CMDADD1 
         Caption         =   "&Add"
         Height          =   495
         Left            =   8040
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin MSComctlLib.ListView LSTITEM 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648384
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   6135
      Left            =   120
      ScaleHeight     =   6075
      ScaleWidth      =   11595
      TabIndex        =   11
      Top             =   240
      Width           =   11655
   End
End
Attribute VB_Name = "FRMREPLACEMENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDADD_Click()
FRMBROWSESTOCKS.Show
AddStocks = "Rep"
End Sub

Private Sub CMDADD1_Click()
AddStocks = "Item"
FRMBROWSESTOCKS.Show
End Sub

Private Sub CMDCLOSE_Click()
Unload Me
End Sub

Private Sub CMDREMOVE_Click()
If LSTITEM.ListItems.Count = 0 Then
        MsgBox "List is Empty there's nothing to Remove ", vbInformation, "Message"
        Exit Sub
    End If
    If MsgBox("Are you sure you want to Remove ?" & " " & _
    LSTITEM.SelectedItem.SubItems(1), vbExclamation + vbYesNo, "Confirm") = vbYes Then
    RemoveOrder LSTITEM.SelectedItem.Text
    End If
End Sub

Private Sub CMDSAVE_Click()
Dim o, s As Long
Dim Query As String


For o = 1 To LSTITEM.ListItems.Count
Query = "INSERT INTO Replacement (stockno,[desc],qty,[date_ret],or_no,clerk) VALUES ('" & LSTITEM.ListItems.Item(o).Text & "','" & LSTITEM.ListItems.Item(o).SubItems(1) & "'," & Val(LSTITEM.ListItems.Item(o).SubItems(2)) & ",#" & Format(Now, "mm/dd/yyyy") & "#,'" & ord & "','" & Mid$(MDIMAIN.StatusBar1.Panels(3).Text, 8, 15) & "');"
DBMain.Execute Query
Next o
 

For o = 1 To LSTITEM.ListItems.Count
UpdateStockQuantity LSTITEM.ListItems.Item(o).SubItems(2), LSTITEM.ListItems.Item(o).Text
Next o
OcashStat = "OK"
ord = ""

Unload Me
End Sub

Private Sub Form_Load()

Me.Caption = Me.Caption & ":" & Format(Now, " HH:MM:SS AMPM")
Call ColForm(Picture1, 217, 211, 213, 125)
Call ColForm(Picture2, 217, 211, 213, 125)
SetFlatList LSTITEM, Me
SetFlatList LSTREP, Me
CMDADD.Enabled = False
CMDREMOVE1.Enabled = False
End Sub
Private Sub RemoveOrder(ByVal Stk As String)
    Dim List As ListItem
    Dim TQty As Long
    Set List = LSTITEM.FindItem(Stk, , , lvwPartial)
    TQty = LSTITEM.SelectedItem.SubItems(2)
    LSTITEM.ListItems.Remove List.Index
    'UpdateStockQuantity TQty, Stk
    
End Sub
Private Sub Form_unload(Cancel As Integer)
Dim x As Long
Dim Query As String
    If Not OcashStat = "OK" Then
    If LSTREP.ListItems.Count <> 0 Then
        
        For x = 1 To LSTREP.ListItems.Count
            Query = "UPDATE stocks SET Quantity = Quantity + " & LSTREP.ListItems(x).SubItems(2) & " WHERE stockno ='" & LSTREP.ListItems(x).Text & "' ;"
            DBMain.Execute Query
        Next x
        End If
        OcashStat = ""
    End If
End Sub
