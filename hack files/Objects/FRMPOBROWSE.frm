VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMPOBROWSE 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse Purchase Order"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "FRMPOBROWSE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7830
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList IMGLIST 
      Left            =   2925
      Top             =   1500
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
            Picture         =   "FRMPOBROWSE.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CMDSELECT 
      Caption         =   "Select"
      Height          =   465
      Left            =   5280
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton CMDCANCEL 
      Caption         =   "Cancel"
      Height          =   465
      Left            =   6495
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin MSComctlLib.ListView LSTPO 
      Height          =   2670
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   4710
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
         Text            =   "Purchase Number"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date Ordered"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Company Name"
         Object.Width           =   6262
      EndProperty
   End
End
Attribute VB_Name = "FRMPOBROWSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CMDCANCEL_Click()
    Unload Me
End Sub

Private Sub CMDSELECT_Click()
    On Error GoTo trapper
    Dim Query As String
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    
    Query = "SELECT pos.ponum,suppliers.suppno,suppliers.company,suppliers.telno FROM pos INNER JOIN suppliers ON pos.suppno = suppliers.suppno WHERE pos.ponum ='" & LSTPO.SelectedItem.Text & "';"
    Set TQR = DBMain.CreateQueryDef("", Query)
    Set TRS = TQR.OpenRecordset()
    With FRMDELIVERY
        .TXTF(0).Text = TRS.Fields(0)
        .TXTF(1).Text = TRS.Fields(1)
        .TXTF(2).Text = TRS.Fields(2)
        .TXTF(3).Text = TRS.Fields(3)
        .TXTF(5).Text = AutoDel
    End With
    Set TQR = Nothing
    Set TRS = Nothing
trapper:
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Query As String
    Dim TRS As DAO.Recordset
    Dim TQR As DAO.QueryDef
    SetFlatList LSTPO, Me
    
    Query = "SELECT DISTINCT pos.ponum,pos.date,suppliers.company FROM pos INNER JOIN suppliers ON pos.suppno = suppliers.suppno WHERE stat = false ;"
    Set TQR = DBMain.CreateQueryDef("", Query)
    Set TRS = TQR.OpenRecordset()
    DisplayRecord LSTPO, TRS, 2, 1
    Set TQR = Nothing
    Set TRS = Nothing
End Sub

Private Sub LSTPO_DblClick()
    CMDSELECT_Click
End Sub
