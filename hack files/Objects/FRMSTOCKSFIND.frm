VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMSTOCKSFIND 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Find Item"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   Icon            =   "FRMSTOCKSFIND.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3795
      Left            =   255
      TabIndex        =   6
      Top             =   525
      Width           =   7065
      Begin MSComctlLib.ListView LSTRESULTS 
         Height          =   2550
         Left            =   15
         TabIndex        =   4
         Top             =   1185
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   4498
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Number "
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description "
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.ComboBox CBOFIELDS 
         Height          =   315
         ItemData        =   "FRMSTOCKSFIND.frx":08CA
         Left            =   1395
         List            =   "FRMSTOCKSFIND.frx":08D4
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   2400
      End
      Begin VB.TextBox TXTVALUE 
         BackColor       =   &H00FFFFFF&
         Height          =   400
         Left            =   1395
         TabIndex        =   0
         Top             =   180
         Width           =   3420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fields"
         Height          =   195
         Left            =   330
         TabIndex        =   8
         Top             =   825
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Find What ?"
         Height          =   195
         Left            =   315
         TabIndex        =   7
         Top             =   285
         Width           =   870
      End
   End
   Begin VB.CommandButton CMDCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   450
      Left            =   7320
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton CMDFIND 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   450
      Left            =   7320
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4275
      Left            =   150
      TabIndex        =   5
      Top             =   135
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   7541
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "          Find          "
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FRMSTOCKSFIND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fields(1) As String

Private Sub CMDCANCEL_Click()
    Unload Me
End Sub

Private Sub CMDFIND_Click()
    Dim List As ListItem
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    LSTRESULTS.ListItems.Clear
    Set TQR = DBMain.CreateQueryDef("", "SELECT stocks.stockno,stocks.description FROM stocks WHERE " & Fields(CBOFIELDS.ListIndex) & " LIKE '" & TXTVALUE.Text & "';")
    Set TRS = TQR.OpenRecordset()
    Do While Not TRS.EOF
        Set List = LSTRESULTS.ListItems.Add(, , TRS.Fields(0))
        List.SubItems(1) = TRS.Fields(1)
        TRS.MoveNext
    Loop
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
   
    CBOFIELDS.ListIndex = 0
    Fields(0) = "StockNo"
    Fields(1) = "Description"
End Sub

Private Sub LSTRESULTS_DblClick()
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim x As Long
    Set TQR = DBMain.CreateQueryDef("", "SELECT * FROM stocks WHERE stockno ='" & LSTRESULTS.SelectedItem.Text & "';")
    Set TRS = TQR.OpenRecordset()
    For x = 0 To 4
        FRMSTOCKS.TXTF(x).Text = TRS.Fields(x)
    Next x
    Unload Me
End Sub



