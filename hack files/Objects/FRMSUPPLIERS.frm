VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMSUPPLIERS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Suppliers MasterFile"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11445
   FillColor       =   &H80000012&
   Icon            =   "FRMSUPPLIERS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList IMGLIST 
      Left            =   5100
      Top             =   3495
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
            Picture         =   "FRMSUPPLIERS.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LSTSUPPS 
      Height          =   3195
      Left            =   600
      TabIndex        =   0
      Top             =   2640
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5636
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Supplier Number"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Company Name"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Telephone #"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "[Supplier Information]"
      Height          =   5775
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   10815
      Begin VB.CommandButton CMDADD 
         Caption         =   "&Add"
         Height          =   800
         Left            =   9360
         Picture         =   "FRMSUPPLIERS.frx":0BEA
         TabIndex        =   12
         Top             =   360
         Width           =   1300
      End
      Begin VB.CommandButton CMDMODIFY 
         Caption         =   "&Modify"
         Height          =   800
         Left            =   9360
         Picture         =   "FRMSUPPLIERS.frx":14B4
         TabIndex        =   11
         Top             =   1200
         Width           =   1300
      End
      Begin VB.CommandButton CMDSAVE 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   800
         Left            =   9360
         Picture         =   "FRMSUPPLIERS.frx":18F6
         TabIndex        =   10
         Top             =   2040
         Width           =   1300
      End
      Begin VB.CommandButton CMDCLOSE 
         Caption         =   "Close"
         Height          =   800
         Left            =   9360
         Picture         =   "FRMSUPPLIERS.frx":1F60
         TabIndex        =   9
         Top             =   2880
         Width           =   1300
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1620
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1155
         Width           =   2955
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1620
         MaxLength       =   30
         TabIndex        =   4
         Top             =   795
         Width           =   4470
      End
      Begin VB.TextBox TXTF 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   390
         Width           =   2385
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3795
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   6694
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "          Suppliers          "
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label LBLCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Telephone # :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1275
         Width           =   1005
      End
      Begin VB.Label LBLCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Company Name :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   795
         Width           =   1215
      End
      Begin VB.Label LBLCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Supplier Number :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   9
      Height          =   6015
      Left            =   240
      Top             =   240
      Width           =   11055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   7
      Height          =   6255
      Left            =   120
      Top             =   120
      Width           =   11295
   End
End
Attribute VB_Name = "FRMSUPPLIERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyType As String

Private Sub CMDADD_Click()
    TXTF(0).Text = AutoSupplier
    TXTF(1).Text = ""
    TXTF(2).Text = ""
    'CMDCANCEL.Enabled = True
    CMDSAVE.Enabled = True
    
    CMDADD.Enabled = False
    CMDMODIFY.Enabled = False
    'CMDDELETE.Enabled = False
    CMDCLOSE.Enabled = False
    
    TXTF(1).Enabled = True
    TXTF(2).Enabled = True
    TXTF(1).SetFocus
    LSTSUPPS.Enabled = False
    MyType = "ADD"
End Sub



Private Sub CMDCLOSE_Click()
'MDIMAIN.STATUSBAR.Panels(2).Text = "Status:"
    Unload Me
End Sub



Private Sub CMDMODIFY_Click()
    If TXTF(0).Text = "" Then
        MsgBox "There is no Record to Modify", vbExclamation, "Message"
        Exit Sub
    End If
    'CMDCANCEL.Enabled = True
    CMDSAVE.Enabled = True
    
    CMDADD.Enabled = False
    CMDMODIFY.Enabled = False
    'CMDDELETE.Enabled = False
    CMDCLOSE.Enabled = False
    TXTF(1).Enabled = True
    TXTF(2).Enabled = True
    TXTF(1).SetFocus
    LSTSUPPS.Enabled = False
    MyType = "EDIT"
End Sub

Private Sub CMDSAVE_Click()
    Dim List As ListItem
    
    Dim Query As String
    If MyType = "ADD" Then
        Query = "INSERT INTO suppliers (suppno,company,telno) VALUES  ('" & TXTF(0).Text & "','" & TXTF(1).Text & "','" & TXTF(2).Text & "');"
        DBMain.Execute Query
        Set List = LSTSUPPS.ListItems.Add(, , TXTF(0).Text, , 1)
        With List
            .SubItems(1) = TXTF(1).Text
            .SubItems(2) = TXTF(2).Text
        End With
        'CMDCANCEL_Click
    ElseIf MyType = "EDIT" Then
        Query = "UPDATE suppliers SET company = '" & TXTF(1).Text & "', telno = '" & TXTF(2).Text & "' WHERE suppno = '" & TXTF(0).Text & "';"
        DBMain.Execute Query
        Set List = LSTSUPPS.FindItem(TXTF(0).Text, , , lvwPartial)
        With List
            .SubItems(1) = TXTF(1).Text
            .SubItems(2) = TXTF(2).Text
        End With
        LSTSUPPS.Enabled = True
        TXTF(1).Enabled = False
        TXTF(2).Enabled = False
        CMDADD.Enabled = True
        'CMDDELETE.Enabled = True
        CMDMODIFY.Enabled = True
        CMDCLOSE.Enabled = True
        CMDSAVE.Enabled = False
        'CMDCANCEL.Enabled = False
    End If
End Sub


Private Sub Form_Load()
    On Error GoTo trapper
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim List As ListItem
    Set TQR = DBMain.CreateQueryDef("", "SELECT * FROM suppliers ORDER BY suppno ASC ")
    Set TRS = TQR.OpenRecordset()
   
    Do While Not TRS.EOF
        Set List = LSTSUPPS.ListItems.Add(, , TRS.Fields(0), , 1)
        List.SubItems(1) = TRS.Fields(1)
        List.SubItems(2) = TRS.Fields(2)
        TRS.MoveNext
    Loop
    With LSTSUPPS.ListItems(1)
        TXTF(0).Text = .Text
        TXTF(1).Text = .SubItems(1)
        TXTF(2).Text = .SubItems(2)
    End With
trapper:

End Sub

Private Sub LSTSUPPS_ItemClick(ByVal Item As MSComctlLib.ListItem)
    TXTF(0).Text = Item.Text
    TXTF(1).Text = Item.SubItems(1)
    TXTF(2).Text = Item.SubItems(2)
    
End Sub

