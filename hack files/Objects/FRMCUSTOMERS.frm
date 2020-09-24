VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMCUSTOMERS 
   BackColor       =   &H8000000A&
   Caption         =   "Customer MasterFile"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ForeColor       =   &H80000006&
   Icon            =   "FRMCUSTOMERS.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   6855
      Left            =   10185
      ScaleHeight     =   6795
      ScaleWidth      =   1515
      TabIndex        =   17
      Top             =   480
      Width           =   1575
      Begin VB.CommandButton CMDADD 
         Caption         =   "&Add"
         Height          =   800
         Left            =   120
         Picture         =   "FRMCUSTOMERS.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   1300
      End
      Begin VB.CommandButton CMDMODIFY 
         Caption         =   "&Modify"
         Height          =   800
         Left            =   120
         Picture         =   "FRMCUSTOMERS.frx":170C
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   960
         Width           =   1300
      End
      Begin VB.CommandButton CMDDELETE 
         Caption         =   "&Delete"
         Height          =   800
         Left            =   120
         Picture         =   "FRMCUSTOMERS.frx":1B4E
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1800
         Width           =   1300
      End
      Begin VB.CommandButton CMDSAVE 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   800
         Left            =   120
         Picture         =   "FRMCUSTOMERS.frx":2418
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2640
         Width           =   1300
      End
      Begin VB.CommandButton CMDCANCEL 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   800
         Left            =   120
         Picture         =   "FRMCUSTOMERS.frx":325A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4800
         Width           =   1300
      End
      Begin VB.CommandButton CMDFIND 
         Caption         =   "&Find"
         Height          =   800
         Left            =   120
         Picture         =   "FRMCUSTOMERS.frx":335C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3360
         Width           =   1300
      End
      Begin VB.CommandButton CMDCLOSE 
         Caption         =   "&Close"
         Height          =   855
         Left            =   120
         Picture         =   "FRMCUSTOMERS.frx":345E
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5760
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Customer Information ]"
      ForeColor       =   &H00C00000&
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   9855
      Begin VB.TextBox TXTF 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   9
         Top             =   360
         Width           =   2580
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   8
         Top             =   840
         Width           =   4410
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1320
         Width           =   4410
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   7800
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   7800
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1800
         Width           =   1965
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1800
         Width           =   4395
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   7800
         MaxLength       =   1
         TabIndex        =   3
         Top             =   840
         Width           =   900
      End
      Begin VB.Label LBLCAPS 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer No:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   960
      End
      Begin VB.Label LBLCAPS 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Firstname :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   765
      End
      Begin VB.Label LBLCAPS 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Lastname :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label LBLCAPS 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Address :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label LBLCAPS 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Telephone No :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   6600
         TabIndex        =   12
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label LBLCAPS 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Credit Limit :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   6840
         TabIndex        =   11
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "MI :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   7440
         TabIndex        =   10
         Top             =   960
         Width           =   270
      End
   End
   Begin MSComctlLib.ImageList IMGLST 
      Left            =   8010
      Top             =   1800
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
            Picture         =   "FRMCUSTOMERS.frx":38A0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LSTCUSTOMERS 
      Height          =   3705
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   6535
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "IMGLST"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Customer Number"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Firstname"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Lastname"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Mi"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Address"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Telephone #"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Credit Limit"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4275
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   7541
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "          Customers          "
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FRMCUSTOMERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyType As String
Dim PassString As String

Private Sub CMDCLOSE_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim x As Long
Call ColForm(Picture1, 217, 211, 213, 125)
FillDatas
    For x = 0 To 5
        TextUpper TXTF(x)
    Next x
    NumberOnly TXTF(6)
  
End Sub

Private Sub CMDADD_Click()
    SetText False
    
    CMDSAVE.Enabled = True
    CMDCANCEL.Enabled = True
    
    LSTCUSTOMERS.Enabled = False
    CMDDELETE.Enabled = False
    CMDCLOSE.Enabled = False
    CMDMODIFY.Enabled = False
    CMDADD.Enabled = False
    CMDFIND.Enabled = False
    cleartext
    TXTF(1).SetFocus
    TXTF(0).Text = AutoCustomerNumber
    MyType = "ADD"
    End Sub

Private Sub CMDCANCEL_Click()
    SetText False
    CMDSAVE.Enabled = False
    CMDCANCEL.Enabled = False
    
    CMDDELETE.Enabled = True
    CMDCLOSE.Enabled = True
    CMDMODIFY.Enabled = True
    CMDADD.Enabled = True
    CMDFIND.Enabled = True
    LSTCUSTOMERS.Enabled = True
    cleartext
    MyType = ""

End Sub



Private Sub CMDDELETE_Click()
    If TXTF(0).Text = "" Then
        MsgBox "Please select a record to be deleted", vbExclamation, "Confirm"
        Exit Sub
    End If
    If MsgBox("Are you sure you want to delete record of " + TXTF(2).Text + ", " + TXTF(1).Text, vbExclamation + vbYesNo, "Deletion confirm") = vbYes Then
        Dim Dellist As ListItem
        MsgBox "All Associated records of " + TXTF(2).Text + ", " + TXTF(1).Text + " Would be deleted Permanently ", vbInformation, "Delete Record"
        DBMain.Execute "DELETE * FROM customers WHERE custno='" + TXTF(0).Text + "';"
        DBMain.Execute "DELETE * FROM Accounts WHERE custno='" + TXTF(0).Text + "';"
        DBMain.Execute "DELETE * FROM Payments WHERE custno='" + TXTF(0).Text + "';"
        Set Dellist = LSTCUSTOMERS.FindItem(TXTF(0).Text, , , lvwPartial)
        LSTCUSTOMERS.ListItems.Remove Dellist.Index
        cleartext
    End If

End Sub

Private Sub CMDFIND_Click()
    FRMFIND.Show vbModal
End Sub

Private Sub CMDMODIFY_Click()
    If TXTF(0).Text = "" Then
        MsgBox "There's no Record to Modify ", vbExclamation, "Confirm"
        Exit Sub
    End If
    SetText True
    LSTCUSTOMERS.Enabled = False
    CMDADD.Enabled = False
    CMDDELETE.Enabled = False
    CMDCLOSE.Enabled = False
    CMDMODIFY.Enabled = False
    CMDFIND.Enabled = False
    
    CMDSAVE.Enabled = True
    CMDCANCEL.Enabled = True
    MyType = "EDIT"

End Sub

Private Sub CMDSAVE_Click()
    On Error GoTo trapper
    Dim TQR As DAO.QueryDef
    
    Dim PassString As String
    Dim List As ListItem
    Dim e As Long
    Dim x As Long
    Dim Flag As Boolean
    Flag = False
    For x = 0 To 6
        If TXTF(x).Text = "" Then Flag = True
    Next x
    If Flag Then
        MsgBox "Please Enter all information to Continue ?", vbInformation, "Confirm"
        TXTF(1).SetFocus
        GoTo x
    End If
    If MyType = "ADD" Then
        PassString = "INSERT INTO customers (CUSTNO,FIRSTNAME,LASTNAME,MIDNAME,ADDRESS,TELNO,CREDITLIMIT) VALUES ('" _
                  & TXTF(0).Text & "','" & TXTF(1) & "','" & TXTF(2) & "','" & TXTF(3) _
                  & "','" & TXTF(4) & "','" & TXTF(5) & "','" & TXTF(6) & "');"
        
        Set TQR = DBMain.CreateQueryDef("", PassString)
        TQR.Execute
        Set List = LSTCUSTOMERS.ListItems.Add(, , TXTF(0).Text, , 1)
        With List
            .SubItems(1) = TXTF(1).Text
            .SubItems(2) = TXTF(2).Text
            .SubItems(3) = TXTF(3).Text
            .SubItems(4) = TXTF(4).Text
            .SubItems(5) = TXTF(5).Text
            .SubItems(6) = TXTF(6).Text
        End With
        CMDCANCEL_Click
    ElseIf MyType = "EDIT" Then
        PassString = "UPDATE customers SET firstname='" & TXTF(1).Text & "', lastname='" & TXTF(2).Text & "', midname='" _
        & TXTF(3).Text & "', address ='" & TXTF(4).Text & "', telno ='" & TXTF(5).Text & "', creditlimit =" & TXTF(6).Text & " WHERE custno='" & TXTF(0).Text & "';"
        DBMain.Execute PassString
        SetText False
        LSTCUSTOMERS.Enabled = True
        CMDADD.Enabled = True
        CMDDELETE.Enabled = True
        CMDCLOSE.Enabled = True
        CMDMODIFY.Enabled = True
        CMDFIND.Enabled = True
        CMDSAVE.Enabled = False
        CMDCANCEL.Enabled = False
        MyType = ""
        Set List = LSTCUSTOMERS.FindItem(TXTF(0).Text, , , lvwPartial)
        With List
            .SubItems(1) = TXTF(1).Text
            .SubItems(2) = TXTF(2).Text
            .SubItems(3) = TXTF(3).Text
            .SubItems(4) = TXTF(4).Text
            .SubItems(5) = TXTF(5).Text
            .SubItems(6) = TXTF(6).Text
        End With
    End If
x:
    Exit Sub
trapper:
    If Err.Number = 3075 Then
        MsgBox "Please input valid data to continue ?", vbExclamation, "Confirm"
        CMDCANCEL_Click
    End If

End Sub

Private Sub Form_Activate()
    DrawBorder 10, Me
End Sub

Private Sub LSTCUSTOMERS_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo trapper
    Dim x As Long
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Set TQR = DBMain.CreateQueryDef("", "SELECT * FROM customers WHERE custno='" & LSTCUSTOMERS.SelectedItem.Text & "';")
    Set TRS = TQR.OpenRecordset()
    For x = 0 To 6
        TXTF(x).Text = TRS.Fields(x)
    Next x

trapper:

End Sub



