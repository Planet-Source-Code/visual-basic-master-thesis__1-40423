VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMPOBROW 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suppliers"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   Icon            =   "FRMPOBROW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8100
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CMDSELECT 
      Caption         =   "Select"
      Default         =   -1  'True
      Height          =   450
      Left            =   5445
      TabIndex        =   2
      Top             =   4005
      Width           =   1215
   End
   Begin VB.CommandButton CMDCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   450
      Left            =   6705
      TabIndex        =   1
      Top             =   4005
      Width           =   1215
   End
   Begin VB.Frame FRABROWSE 
      BackColor       =   &H8000000A&
      Caption         =   "[ Available Suppliers ]"
      ForeColor       =   &H00000000&
      Height          =   3690
      Left            =   165
      TabIndex        =   0
      Top             =   180
      Width           =   7770
      Begin MSComctlLib.ImageList IMGLIST 
         Left            =   3765
         Top             =   2055
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
               Picture         =   "FRMPOBROW.frx":000C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView LSTSUPPS 
         Height          =   3195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   7515
         _ExtentX        =   13256
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
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Telephone #"
            Object.Width           =   3528
         EndProperty
      End
   End
End
Attribute VB_Name = "FRMPOBROW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDCANCEL_Click()
    Unload Me
End Sub

Private Sub CMDSELECT_Click()
    With LSTSUPPS
        FRMPO.TXTF(0).Text = .SelectedItem.Text
        FRMPO.TXTF(1).Text = .SelectedItem.SubItems(1)
        FRMPO.TXTF(2).Text = .SelectedItem.SubItems(2)
    End With
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        CMDCANCEL_Click
    End If
End Sub

Private Sub Form_Load()
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Set TQR = DBMain.QueryDefs("ALLSupplier")
    Set TRS = TQR.OpenRecordset()
    SetFlatList LSTSUPPS, Me
    DisplayRecord LSTSUPPS, TRS, 2, 1
End Sub

Private Sub LSTSUPPS_DblClick()
    CMDSELECT_Click
End Sub
