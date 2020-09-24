VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMVIEWCUST 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Statement of Accounts"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   Icon            =   "FRMVIEWCUST.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7740
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMVIEWCUST.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FRAVIEW 
      BackColor       =   &H8000000A&
      Caption         =   "[ Customers ]"
      ForeColor       =   &H00000000&
      Height          =   3105
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   7515
      Begin MSComctlLib.ListView LSTCUSTOMERS 
         Height          =   2775
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Customer Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Firstname"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Lastname"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Address"
            Object.Width           =   4304
         EndProperty
      End
   End
   Begin VB.CommandButton CMDSELECT 
      Caption         =   "Select"
      Height          =   420
      Left            =   5205
      TabIndex        =   1
      Top             =   3315
      Width           =   1200
   End
   Begin VB.CommandButton CMDCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   6435
      TabIndex        =   0
      Top             =   3315
      Width           =   1200
   End
End
Attribute VB_Name = "FRMVIEWCUST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDCANCEL_Click()
    Unload FRMVIEWCUST
    
End Sub

Private Sub CMDSELECT_Click()
    With FRMPRINTCUSTOMER
        .TXTF(0).Text = LSTCUSTOMERS.SelectedItem.Text
        .TXTF(2).Text = LSTCUSTOMERS.SelectedItem.SubItems(1)
        .TXTF(1).Text = LSTCUSTOMERS.SelectedItem.SubItems(2)
        .TXTF(3).Text = LSTCUSTOMERS.SelectedItem.SubItems(3)
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    Dim TRS As DAO.Recordset
    Dim TQR As DAO.QueryDef
    Dim TList As ListItem
    
    SetFlatList LSTCUSTOMERS, Me
    Set TQR = DBMain.QueryDefs("CView")
    Set TRS = TQR.OpenRecordset()
    Do While Not TRS.EOF
        Set TList = LSTCUSTOMERS.ListItems.Add(, , TRS.Fields(0), , 1)
        With TList
            .SubItems(2) = TRS.Fields(1)
            .SubItems(1) = TRS.Fields(2)
            .SubItems(3) = TRS.Fields(3)
        End With
        TRS.MoveNext
    Loop
End Sub

Private Sub LSTCUSTOMERS_DblClick()
    CMDSELECT_Click
End Sub
