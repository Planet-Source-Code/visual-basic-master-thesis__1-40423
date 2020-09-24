VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMFIND 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   Icon            =   "FRMFIND.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8655
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3795
      Left            =   225
      TabIndex        =   6
      Top             =   540
      Width           =   8160
      Begin MSComctlLib.ListView LSTRESULTS 
         Height          =   2115
         Left            =   90
         TabIndex        =   4
         Top             =   1590
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   3731
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Customer Number"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Lastname"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Firstname"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.CommandButton CMDCANCEL 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   435
         Left            =   6720
         TabIndex        =   3
         Top             =   840
         Width           =   1305
      End
      Begin VB.CommandButton CMDSEARCH 
         Caption         =   "Search"
         Default         =   -1  'True
         Height          =   435
         Left            =   6720
         TabIndex        =   2
         Top             =   285
         Width           =   1305
      End
      Begin VB.TextBox TXTVALUE 
         BackColor       =   &H00FFFFFF&
         Height          =   400
         Left            =   930
         TabIndex        =   0
         Top             =   300
         Width           =   3840
      End
      Begin VB.ComboBox CBOFIELDS 
         Height          =   315
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   870
         Width           =   2235
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         Height          =   195
         Left            =   225
         TabIndex        =   9
         Top             =   420
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fields"
         Height          =   195
         Left            =   225
         TabIndex        =   8
         Top             =   930
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Search Results"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   1335
         Width           =   1080
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4275
      Left            =   135
      TabIndex        =   5
      Top             =   135
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   7541
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "     Search     "
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FRMFIND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fields(5) As String

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub CMDCANCEL_Click()
    Unload Me
End Sub

Private Sub CMDSEARCH_Click()
    On Error GoTo trapper
    Dim TQR As DAO.QueryDef
    Dim TList As ListItem
    Dim TRS As DAO.Recordset
    Dim StrQry As String
    
    If CBOFIELDS.ListIndex = 5 Then
        StrQry = "SELECT custno,lastname,firstname FROM customers WHERE " & Fields(CBOFIELDS.ListIndex) & " LIKE " & TXTVALUE.Text & ";"
    Else
        StrQry = "SELECT custno,lastname,firstname FROM customers WHERE " & Fields(CBOFIELDS.ListIndex) & " LIKE '" & TXTVALUE.Text & "';"
    End If
    
    Set TQR = DBMain.CreateQueryDef("", StrQry)
    Set TRS = TQR.OpenRecordset()
    LSTRESULTS.ListItems.Clear
    Do While Not TRS.EOF
        Set TList = LSTRESULTS.ListItems.Add(, , TRS.Fields(0))
        TList.SubItems(1) = TRS.Fields(1)
        TList.SubItems(2) = TRS.Fields(2)
        TRS.MoveNext
    Loop
    Exit Sub
trapper:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    SetFlatList LSTRESULTS, FRMFIND
    With CBOFIELDS
        .AddItem "Customer Number"
        .AddItem "Lastname"
        .AddItem "Firstname"
        .AddItem "Address"
        .AddItem "Telephone #"
        .AddItem "Credit Limit"
    End With
    Fields(0) = "Custno"
    Fields(1) = "Lastname"
    Fields(2) = "firstname"
    Fields(3) = "Address"
    Fields(4) = "Telno"
    Fields(5) = "Creditlimit"
    CBOFIELDS.ListIndex = 0

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim x As Long
    Set TQR = DBMain.CreateQueryDef("", "SELECT * FROM customers WHERE custno='" & LSTRESULTS.SelectedItem.Text & "';")
    Set TRS = TQR.OpenRecordset()
    With FRMCUSTOMERS.TXTF
    For x = 0 To 6
        .Item(x).Text = TRS.Fields(x)
    Next x
    End With
    Unload Me

End Sub


Private Sub LSTRESULTS_DblClick()
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim x As Long
    Set TQR = DBMain.CreateQueryDef("", "SELECT * FROM customers WHERE custno='" & LSTRESULTS.SelectedItem.Text & "';")
    Set TRS = TQR.OpenRecordset()
    With FRMCUSTOMERS.TXTF
    For x = 0 To 6
        .Item(x).Text = TRS.Fields(x)
    Next x
    End With
    Unload Me

End Sub

