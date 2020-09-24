VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMADDORDERS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Order Form"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11595
   ControlBox      =   0   'False
   Icon            =   "FRMADDORDERS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "[ Available Stocks ]"
      Height          =   2370
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   10935
      Begin VB.CommandButton CMDADD 
         Caption         =   "&Add"
         Height          =   465
         Left            =   9570
         TabIndex        =   6
         Top             =   330
         Width           =   1230
      End
      Begin MSComctlLib.ImageList IMGLIST 
         Left            =   600
         Top             =   1080
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
               Picture         =   "FRMADDORDERS.frx":000C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView LSTASTOCKS 
         Height          =   1965
         Left            =   90
         TabIndex        =   7
         Top             =   315
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   3466
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "[ Orders ]"
      ForeColor       =   &H00000000&
      Height          =   2940
      Left            =   360
      TabIndex        =   0
      Top             =   2880
      Width           =   10965
      Begin VB.CommandButton CMDREMOVE 
         Caption         =   "&Remove"
         Height          =   465
         Left            =   9645
         TabIndex        =   3
         Top             =   375
         Width           =   1230
      End
      Begin VB.CommandButton CMDORDER 
         Caption         =   "&Order"
         Height          =   465
         Left            =   9660
         TabIndex        =   2
         Top             =   1560
         Width           =   1230
      End
      Begin VB.CommandButton CMDCLOSE 
         Caption         =   "Close"
         Height          =   465
         Left            =   9660
         TabIndex        =   1
         Top             =   2160
         Width           =   1230
      End
      Begin MSComctlLib.ListView LSTORDERS 
         Height          =   2475
         Left            =   90
         TabIndex        =   4
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4366
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
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   360
      ScaleHeight     =   1275
      ScaleWidth      =   10875
      TabIndex        =   9
      Top             =   6000
      Width           =   10935
      Begin VB.TextBox TXTDATE 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label LBLITEMS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6600
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label LBLTOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2160
         TabIndex        =   16
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label LBLCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Total Amount Due:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1920
      End
      Begin VB.Label LBLORDNO 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "ORDER NUMBER : 0-00001"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   2835
      End
      Begin VB.Label LBLORDDATE 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "ORDER DATE :12/12/2000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   3600
         TabIndex        =   13
         Top             =   120
         Width           =   2730
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Due Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   6840
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         Caption         =   "No. Of  Item:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5160
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   5130
      TabIndex        =   8
      Top             =   3180
      Width           =   1215
   End
End
Attribute VB_Name = "FRMADDORDERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDADD_Click()
    Dim List As ListItem
    If LSTASTOCKS.ListItems.Count = 0 Then
        MsgBox "No Current stocks to ADD", vbInformation, "Message"
        Exit Sub
    End If
    AddOrder LSTASTOCKS.SelectedItem.Text, GetQuantity
End Sub

Private Sub CMDCLOSE_Click()
    OCreditStat = "BAD"
    
    Unload Me
End Sub

Private Sub CMDORDER_Click()

    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim Query As String
    Dim Count As Long
    
    'check if the customer exceed his/her creditlimit
    '---------------------------------------------------'
    If CheckCreditLimit(Me.Tag) Then
    MsgBox "Customer has rearch creditLimit", vbOKOnly + vbInformation, "Order"
    Exit Sub
    End If
   '-----------------------------------------------------'
    If MsgBox("Are you sure you want to Order the Following Item(s) ", vbExclamation + vbYesNo, "Confirm Order") = vbYes Then
        For Count = 1 To LSTORDERS.ListItems.Count
            With LSTORDERS.ListItems
            Query = "INSERT INTO orders (ordno,custno,stockno,invno,amount,quantity) VALUES ('" & LBLORDNO.Tag & "','" & Me.Tag & "','" & .Item(Count).Text & "','DONT'," & .Item(Count).SubItems(3) & "," & .Item(Count).SubItems(2) & ");"
            End With
            DBMain.Execute Query
        Next Count
        Query = "INSERT INTO sales (ordno,amount,[date],type,paytype) VALUES ('" & LBLORDNO.Tag & "'," & OrdAmount & ",#" & Format(Now, "mm/dd/yyyy") & "#,'CREDIT','NONE');"
        DBMain.Execute Query
        Query = "INSERT INTO accounts (custno,ordno,amount,[date],balance,[due_date]) VALUES ('" & Me.Tag & "','" & LBLORDNO.Tag & "'," & OrdAmount & ",#" & Format(Now, "mm/dd/yyyy") & "#," & OrdAmount & ",#" & TXTDATE.Text & "#);"
        DBMain.Execute Query
        MsgBox "The items you ordered was Successfully added to your account !", vbInformation, "Message"
        OCreditStat = "OK"
        FRMOCREDIT.LSTCUSTORD.ListItems.Clear
        
        Change Me.Tag
        DueAmount Me.Tag
        Unload Me
    End If
End Sub

Private Sub CMDREMOVE_Click()
   Dim x As Long
    If LSTORDERS.ListItems.Count = 0 Then
        MsgBox "There's no Orders to Remove ", vbInformation, "Message"
        Exit Sub
    End If
    If MsgBox("Are you sure you want to Remove ? " & LSTORDERS.SelectedItem.SubItems(1), vbExclamation + vbYesNo, "Confirm") = vbYes Then
        RemoveOrder LSTORDERS.SelectedItem.Text, LSTORDERS.SelectedItem.SubItems(2)
  End If
End Sub

Private Sub Form_Activate()
   TXTDATE.Text = Format(Now + 15, "mm/dd/yyyy")
    DrawBorder 50, Me

End Sub

Private Sub Form_Load()
    Call ColForm(Picture1, 217, 211, 213, 125)
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Set TQR = DBMain.QueryDefs("AllStocks")
    Set TRS = TQR.OpenRecordset()
    DisplayRecord LSTASTOCKS, TRS, 4, 1
    SetFlatList LSTORDERS, Me
    SetFlatList LSTASTOCKS, Me
    LBLITEMS.Caption = ""
    LBLITEMS.Caption = Val(LBLITEMS.Caption)
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
        MsgBox "Stock Reach Re-Order Level", vbCritical, "Warning !"
        Exit Sub
    End If
    ' check if stocks overloaded
    If Qty > TRS.Fields("Quantity") Then
        MsgBox "Insuffecient Stock Quantity. ", vbCritical, "Warning !"
        Exit Sub
    End If
    
    ' use find method to see if the user already entered this stocks
    ' if stocks is already entered then append the quantity to the list
    ' else add a new list to lstorders
    
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
    LBLTOTAL.Caption = Format(CStr(OrdAmount), "###,###,##0.00")
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
  If Not OCreditStat = "OK" Then
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

Private Sub RemoveOrder(ByVal Number As String, ByVal Qty As Long)
    Dim List As ListItem
    Dim TQty As Long
    Set List = LSTORDERS.FindItem(Number, , , lvwPartial)
    TQty = LSTORDERS.SelectedItem.SubItems(2)
    LSTORDERS.ListItems.Remove List.Index
    Set List = LSTASTOCKS.FindItem(Number, , , lvwPartial)
    List.SubItems(3) = TQty + List.SubItems(3)
    UpdateStockQuantity List.SubItems(3), Number
    LBLITEMS.Caption = Val(LBLITEMS.Caption) - TQty
    OrdAmount = Val(ComputeAmount)
    LBLTOTAL.Caption = OrdAmount
End Sub

Private Sub LSTASTOCKS_DblClick()
    CMDADD_Click
End Sub

