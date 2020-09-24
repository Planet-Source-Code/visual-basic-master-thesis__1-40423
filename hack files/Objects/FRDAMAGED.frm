VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMDAMAGED 
   Caption         =   "Damaged Item(s)"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   11520
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "List of Item(s)"
      Height          =   3135
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   10695
      Begin VB.TextBox TXTSEARCH 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   9015
      End
      Begin VB.CommandButton CMDADD 
         Caption         =   "&Add"
         Height          =   495
         Left            =   9240
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin MSComctlLib.ListView LSTASTOCKS 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4048
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
            Text            =   "Item Number"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   8114
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "R_Order_Lvl"
            Object.Width           =   2293
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Damaged Item(s)"
      Height          =   3015
      Left            =   360
      TabIndex        =   2
      Top             =   3840
      Width           =   10695
      Begin VB.CommandButton CMDSAVE 
         Appearance      =   0  'Flat
         Caption         =   "&Save"
         Height          =   495
         Left            =   9240
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Close"
         Height          =   495
         Left            =   9240
         TabIndex        =   6
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton CMDREMOVE 
         Caption         =   "&Remove"
         Height          =   495
         Left            =   9240
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin MSComctlLib.ListView LSTDAMAGED 
         Height          =   2535
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date_Return"
            Object.Width           =   2469
         EndProperty
      End
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   8
      Height          =   6855
      Left            =   240
      Top             =   480
      Width           =   11055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   6
      Height          =   7095
      Left            =   120
      Top             =   360
      Width           =   11295
   End
   Begin VB.Label Label1 
      Caption         =   "Total Number Of Records:"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   6960
      Width           =   3375
   End
End
Attribute VB_Name = "FRMDAMAGED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDADD_Click()
 Dim List As ListItem
    If LSTASTOCKS.ListItems.Count = 0 Then
        MsgBox "No Current item to ADD", vbInformation, "Message"
        Exit Sub
    End If
    AddOrder LSTASTOCKS.SelectedItem.Text, GetQuantity
End Sub


Private Sub RemoveOrder(ByVal StkNo As String, ByVal Qty As Long)
    Dim List As ListItem
    Dim TQty As Long
    Set List = LSTDAMAGED.FindItem(StkNo, , , lvwPartial)
    TQty = LSTDAMAGED.SelectedItem.SubItems(2)
    LSTDAMAGED.ListItems.Remove List.Index
    Set List = LSTASTOCKS.FindItem(StkNo, , , lvwPartial)
    List.SubItems(2) = TQty + List.SubItems(2)
    UpdateStocksQuantity List.SubItems(2), StkNo
End Sub

Private Sub CMDREMOVE_Click()
   If LSTDAMAGED.ListItems.Count = 0 Then
        MsgBox "There's nothing to Remove ", vbInformation, "Message"
        Exit Sub
    End If
    If MsgBox("Are you sure you want to Remove ? " & LSTDAMAGED.SelectedItem.SubItems(1), vbExclamation + vbYesNo, "Confirm") = vbYes Then
        RemoveOrder LSTDAMAGED.SelectedItem.Text, LSTDAMAGED.SelectedItem.SubItems(2)
    End If
End Sub

Private Sub CMDSAVE_Click()
Dim Counter As Long
Dim Query As String
Dim Description As String


If LSTDAMAGED.ListItems.Count = 0 Then Exit Sub
For Counter = 1 To LSTDAMAGED.ListItems.Count
Query = "INSERT INTO DAMAGED (Stockno,[Desc],Qty,[Date_ret],Replaced,Ord_no,clerk) VALUES ('" & LSTDAMAGED.ListItems.Item(Counter).Text & "','" & LSTDAMAGED.ListItems.Item(Counter).SubItems(1) & "'," & LSTDAMAGED.ListItems.Item(Counter).SubItems(2) & ",#" & LSTDAMAGED.ListItems.Item(Counter).SubItems(3) & "#,'False','" & ord & "','NONE');"
DBMain.Execute Query
Next Counter

MsgBox "Damaged Item was successfully recorded", vbOKOnly + vbInformation, "Damaged"
LSTDAMAGED.ListItems.Clear
Status = "OK"
ord = ""
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
ShowStocks
End Sub
Private Sub Form_unload(Cancel As Integer)
    If Not Status = "OK" Then
    If LSTDAMAGED.ListItems.Count <> 0 Then
        Dim x As Long
        Dim Query As String
        For x = 1 To LSTDAMAGED.ListItems.Count
            Query = "UPDATE stocks SET Quantity = Quantity + " & LSTDAMAGED.ListItems(x).SubItems(2) & " WHERE stockno ='" & LSTDAMAGED.ListItems(x).Text & "' ;"
            DBMain.Execute Query
        Next x
        End If
        Status = ""
     
    End If
End Sub
Private Sub AddOrder(ByVal StkNo As String, ByVal Qty As Long)
    Dim List As ListItem
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    
    Dim Amount As Double
    Dim NQty As Long
    
    If Qty = 0 Then Exit Sub
    Set TQR = DBMain.CreateQueryDef("", "SELECT * FROM stocks WHERE stockno ='" & StkNo & "';")
    Set TRS = TQR.OpenRecordset()
    ' check if it reach re-order level
    If TRS.Fields("Quantity") <= TRS.Fields("ReOrder") Then
        MsgBox "Item Reach Re-Order Level", vbCritical, "Warning !"
        Exit Sub
    End If
    ' check if stocks overloaded
    If Qty > TRS.Fields("Quantity") Then
        MsgBox "Insuffecient Item Quantity. ", vbCritical, "Warning !"
        Exit Sub
    End If
    
    ' use find method to see if the user already entered this stocks
    ' if stocks is already entered then append the quantity to the list
    ' else add a new list to lstorders
    
    Set List = LSTDAMAGED.FindItem(StkNo, , , lvwPartial)
    'Amount = Qty * TRS.Fields("UnitPrice")
    If List Is Nothing Then
        Set List = LSTDAMAGED.ListItems.Add(, , TRS("stockno"), , 0)
        With List
            .SubItems(1) = TRS.Fields(1)
            .SubItems(2) = Qty
            .SubItems(3) = Format(Now, "mm/dd/yyyy")
            
        End With
    Else
        With List
            .SubItems(1) = TRS.Fields(1)
            .SubItems(2) = Qty
            .SubItems(3) = Format(Now, "mm/dd/yyyy")
            
        End With
    End If
    DBMain.Execute "UPDATE stocks SET quantity = quantity - " & Str(Qty) & " WHERE stockno ='" & StkNo & "';"
    Set List = LSTASTOCKS.FindItem(StkNo, , , lvwPartial)
    List.SubItems(2) = TRS.Fields("quantity")
    TXTSEARCH.SelStart = 0
    TXTSEARCH.SelLength = Len(TXTSEARCH.Text)
    TXTSEARCH.SetFocus
End Sub
Public Sub Search1(Lvw As ListView, sFind, Mytextbox As TextBox)
Dim Lvfindtm As ListItem
Dim strTemp As String
Dim s As Long
Set Lvfindtm = Lvw.FindItem(sFind, lvwSubItem, , lvwPartial)
If Not Lvfindtm Is Nothing Then
Lvfindtm.EnsureVisible
Lvfindtm.Selected = True
End If

End Sub

Private Sub TXTSEARCH_Change()
Search1 LSTASTOCKS, Trim(TXTSEARCH.Text), TXTSEARCH
End Sub

Private Sub TXTSEARCH_KeyPress(KeyAscii As Integer)
Alpha keysacii
If KeyAscii = 13 Then
CMDADD.Value = True
End If
End Sub
