VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMOCASH 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "Customer Order Form"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "FRMOCASH.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "----[ Order Form ]"
      ForeColor       =   &H00C00000&
      Height          =   7695
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11415
      Begin VB.TextBox TXTTOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   300
         Left            =   1560
         TabIndex        =   11
         Text            =   "0.0"
         Top             =   6840
         Width           =   4455
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "[ Available Item ]"
         ForeColor       =   &H00000000&
         Height          =   3450
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   11085
         Begin VB.TextBox TXTSEARCH 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   9375
         End
         Begin VB.CommandButton CMDADD 
            Caption         =   "&Add"
            Height          =   465
            Left            =   9600
            TabIndex        =   7
            Top             =   720
            Width           =   1230
         End
         Begin MSComctlLib.ImageList IMGLIST 
            Left            =   720
            Top             =   1920
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
                  Picture         =   "FRMOCASH.frx":000C
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView LSTASTOCKS 
            Height          =   2085
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   3678
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
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
               Text            =   "Item Number "
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description "
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Unit Price"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Quantity "
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Re - Order level"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Item Description"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "[ Orders ]"
         ForeColor       =   &H80000006&
         Height          =   2820
         Left            =   240
         TabIndex        =   1
         Top             =   3720
         Width           =   11085
         Begin VB.CommandButton CMDREMOVE 
            Caption         =   "&Remove"
            Height          =   465
            Left            =   9720
            TabIndex        =   4
            Top             =   240
            Width           =   1230
         End
         Begin VB.CommandButton CMDORDER 
            Caption         =   "&Order"
            Height          =   465
            Left            =   9720
            TabIndex        =   3
            Top             =   1560
            Width           =   1230
         End
         Begin VB.CommandButton CMDCLOSE 
            Caption         =   "Close"
            Height          =   465
            Left            =   9720
            TabIndex        =   2
            Top             =   2160
            Width           =   1230
         End
         Begin MSComctlLib.ListView LSTORDERS 
            Height          =   2475
            Left            =   90
            TabIndex        =   5
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   4366
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "IMGLIST"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Item Number"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description "
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Quantity "
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Amount"
               Object.Width           =   3528
            EndProperty
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "Total Amount:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   6840
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000A&
         Caption         =   "No. Of Item:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6240
         TabIndex        =   13
         Top             =   6840
         Width           =   1335
      End
      Begin VB.Label LBLITEM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7680
         TabIndex        =   12
         Top             =   6840
         Width           =   1575
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   8
      Height          =   7935
      Left            =   120
      Top             =   120
      Width           =   11655
   End
End
Attribute VB_Name = "FRMOCASH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDADD_Click()
    Dim List As ListItem
    If LSTASTOCKS.ListItems.Count = 0 Then
        MsgBox "No Current item to Add", vbInformation, "Message"
        Exit Sub
    End If
    AddOrder LSTASTOCKS.SelectedItem.Text, GetQuantity
End Sub

Private Sub CMDCLOSE_Click()
    Unload Me

End Sub

Private Sub CMDORDER_Click()
    'On Error Resume Next
    If LSTORDERS.ListItems.Count = 0 Then
        MsgBox " No Current order(s) found in the Form", vbExclamation, "Message"
        Exit Sub
    End If
    If MsgBox("Are you sure you want to Order this Items ", vbExclamation + vbYesNo, "Confirm Order") = vbYes Then
        Dim List As ListItem
        Dim x As Long
        For x = 1 To LSTORDERS.ListItems.Count
            Set List = FRMCASHPAYMENT.LSTORD.ListItems.Add(, , LSTORDERS.ListItems(x).Text, , 1)
            With List
                .SubItems(1) = LSTORDERS.ListItems(x).SubItems(1)
                .SubItems(2) = LSTORDERS.ListItems(x).SubItems(2)
                .SubItems(3) = Format(LSTORDERS.ListItems(x).SubItems(3), "###,###,##0.00")
            End With
          Next x
               
               Set List = FRMCASHPAYMENT.LSTORD.ListItems.Add(, , "Total Amount Due:", , 0)
                
                With List
                .SubItems(3) = Format(ComputeAmount, "###,###,##0.00")
                .ListSubItems(3).ForeColor = vbRed
        End With
                Set List = FRMCASHPAYMENT.LSTORD.ListItems.Add(, , "Total Number of Items:", , 0)
                With List
                .SubItems(3) = LBLITEM.Caption
       
        End With
        With FRMCASHPAYMENT
        .LBLMESS(0).Caption = "Order Number : " + AutoOrder
         '.LBLMESS(2).Caption = "Total Amount : " + Format(ComputeAmount, "###,###,##0.00")
         .LBLMESS(0).Tag = AutoOrder
         .LBLMESS(1).Caption = "Date Ordered : " + Format(Now, "MM/DD/YYYY")
         '.LBLNITEMS.Caption = "No of Items Purchase:" + LBLITEM.Caption
         '.LBLNITEMS.Tag = LBLITEM.Caption
      End With
       FRMCASHPAYMENT.Show vbModal
    End If
End Sub

Private Sub CMDREMOVE_Click()
    If LSTORDERS.ListItems.Count = 0 Then
        MsgBox "There's no Orders to Remove ", vbInformation, "Message"
        Exit Sub
    End If
    If MsgBox("Are you sure you want to Remove ? " & LSTORDERS.SelectedItem.SubItems(1), vbExclamation + vbYesNo, "Confirm") = vbYes Then
        RemoveOrder LSTORDERS.SelectedItem.Text, LSTORDERS.SelectedItem.SubItems(2)
    End If
End Sub


Private Sub Form_Load()
TXTTOTAL.Text = "0.00"
LBLITEM.Caption = 0
   DisplayStocks
End Sub

Public Sub DisplayStocks()
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim List As ListItem
    
    Set TQR = DBMain.QueryDefs("AllStocks")
    Set TRS = TQR.OpenRecordset()
    Do While Not TRS.EOF
        Set List = LSTASTOCKS.ListItems.Add(, , TRS.Fields(0), , 1)
        With List
            .SubItems(1) = TRS.Fields(1)
            .SubItems(2) = TRS.Fields(2)
            .SubItems(3) = TRS.Fields(3)
            .SubItems(4) = TRS.Fields(4)
        End With
        TRS.MoveNext
    Loop
End Sub

Private Sub RemoveOrder(ByVal Number As String, ByVal Qty As Long)
    Dim List As ListItem
    Dim TQty As Long
    Set List = LSTORDERS.FindItem(Number, , , lvwPartial)
    Number = LSTORDERS.SelectedItem.Text
    TQty = LSTORDERS.SelectedItem.SubItems(2)
    DBMain.Execute "UPDATE stocks SET quantity = quantity + " & Val(LSTORDERS.SelectedItem.SubItems(2)) & " WHERE stockno ='" & LSTORDERS.SelectedItem.Text & "';"
    LSTORDERS.ListItems.Remove List.Index
    Set List = LSTASTOCKS.FindItem(Number, , , lvwPartial)
    List.SubItems(3) = List.SubItems(3) + TQty
    LBLITEM.Caption = LBLITEM.Caption - TQty
    OrdAmount = Val(ComputeAmount)
    TXTTOTAL.Text = OrdAmount
End Sub


Private Sub AddOrder(ByVal Number As String, ByVal Qty As Long)
    Dim List As ListItem
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    
    Dim Amount As Double
    Dim NQty As Long
    
    If Qty = 0 Then Exit Sub
    Set TQR = DBMain.CreateQueryDef("", "SELECT * FROM stocks WHERE stockno ='" & Number & "';")
    Set TRS = TQR.OpenRecordset()
    ' check if it reach re-order level
    If TRS.Fields("Quantity") <= TRS.Fields("ReOrder") Then
        MsgBox "Item Reach Re-Order Level", vbCritical, "Warning !"
       LBLITEM.Caption = Val(LBLITEM.Caption) - Qty
       Exit Sub
    End If
    ' check if stocks overloaded
    If Qty > TRS.Fields("Quantity") Then
        MsgBox "Insuffecient Item Quantity. ", vbCritical, "Warning !"
        LBLITEM.Caption = Val(LBLITEM.Caption) - Qty
        Exit Sub
    End If
    
    ' use find method to see if the user already entered this stocks
    ' if stocks is already entered then append the quantity to the list
    ' else add a new list to lstorders yeah
    
    Set List = LSTORDERS.FindItem(Number, , , lvwPartial)
    Amount = Qty * TRS.Fields("UnitPrice")
    If List Is Nothing Then
        Set List = LSTORDERS.ListItems.Add(, , TRS("stockno"), , 1)
        With List
            .SubItems(1) = TRS.Fields(1)
            .SubItems(2) = Qty
            .SubItems(3) = Format(Str(Amount), "0.00")
        End With
    Else
        With List
            .SubItems(1) = TRS.Fields(1)
            .SubItems(2) = .SubItems(2) + Qty
            .SubItems(3) = Format(Str(Amount + .SubItems(3)), "0.00")
        End With
    End If
    DBMain.Execute "UPDATE stocks SET quantity = quantity - " & Str(Qty) & " WHERE stockno ='" & Number & "';"
    Set List = LSTASTOCKS.FindItem(Number, , , lvwPartial)
    List.SubItems(3) = TRS.Fields("quantity")
    OrdAmount = Val(ComputeAmount)
    TXTTOTAL.Text = Format(CStr(OrdAmount), "###,###,##0.00")
    TXTSEARCH.SelStart = 0
    TXTSEARCH.SelLength = Len(TXTSEARCH.Text)
    TXTSEARCH.SetFocus

End Sub

Function ComputeAmount() As String
    Dim x As Long
    Dim Total As Double
    For x = 1 To LSTORDERS.ListItems.Count
        Total = Total + Val(LSTORDERS.ListItems(x).SubItems(3))
    Next x
    ComputeAmount = CStr(Total)
End Function
Private Sub Form_unload(Cancel As Integer)
    If Not OcashStat = "OK" Then
    If LSTORDERS.ListItems.Count <> 0 Then
        Dim x As Long
        Dim Query As String
        For x = 1 To LSTORDERS.ListItems.Count
            Query = "UPDATE stocks SET Quantity = Quantity + " & LSTORDERS.ListItems(x).SubItems(2) & " WHERE stockno ='" & LSTORDERS.ListItems(x).Text & "' ;"
            DBMain.Execute Query
        Next x
        End If
        OcashStat = ""
     
    End If
End Sub

Private Sub LSTASTOCKS_DblClick()
    CMDADD_Click
End Sub

Private Sub LSTORDERS_Click()
    Dim x As Long
    For x = 1 To LSTORDERS.ListItems.Count
        Debug.Print "UPDATE stocks SET Quantity = Quantity + " & LSTORDERS.ListItems(x).SubItems(2) & ";"
    Next x
End Sub

Public Sub Search1(Lvw As ListView, sFind)
Dim Lvfindtm As ListItem
Dim strTemp As String
Set Lvfindtm = Lvw.FindItem(sFind, lvwSubItem, , lvwPartial)
If Not Lvfindtm Is Nothing Then
Lvfindtm.EnsureVisible
Lvfindtm.Selected = True
End If

End Sub
Private Sub TXTSEARCH_Change()
Search1 LSTASTOCKS, Trim(TXTSEARCH.Text)
End Sub

Private Sub TXTSEARCH_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CMDADD.Value = True
End If
End Sub
