VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMBROWSE 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Browse By Suppliers"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   Icon            =   "FRMBROWSE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   4440
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
            Picture         =   "FRMBROWSE.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CMDOK 
      Height          =   500
      Left            =   240
      Picture         =   "FRMBROWSE.frx":0BEE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1400
   End
   Begin VB.CommandButton CMDCANCEL 
      Height          =   500
      Left            =   1800
      Picture         =   "FRMBROWSE.frx":0CBD
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1400
   End
   Begin VB.Frame FRABROWSEPARENT 
      BorderStyle     =   0  'None
      Height          =   4170
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9465
      Begin VB.Frame FRABROWSE 
         Caption         =   "[ Browse ]"
         ForeColor       =   &H00FF0000&
         Height          =   3960
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   9360
         Begin MSComctlLib.ListView LSTCLIST 
            Height          =   3615
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   6376
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Supplier No#"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Company"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Tel No#"
               Object.Width           =   4410
            EndProperty
         End
      End
   End
   Begin MSComctlLib.ImageList IMGLIST 
      Left            =   255
      Top             =   5205
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
            Picture         =   "FRMBROWSE.frx":0DBD
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FRMBROWSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CMDCANCEL_Click()
    Unload Me
End Sub

Private Sub CMDOK_Click()
    Dim Customer As String
    FRMDELIVERY.TXTF(1).Text = LSTCLIST.SelectedItem.Text
    FRMDELIVERY.TXTF(2).Text = LSTCLIST.SelectedItem.SubItems(1)
    FRMDELIVERY.TXTF(3).Text = LSTCLIST.SelectedItem.SubItems(2)
     FRMDELIVERY.TXTF(4).Text = Format(Now, "mm/dd/yyyy")
     FRMDELIVERY.TXTF(5).Text = AutoDel
    
    Unload Me
End Sub

Private Sub Form_Load()
    
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim List As ListItem
    
    Set TQR = DBMain.QueryDefs("Browse")
    Set TRS = TQR.OpenRecordset()
    With LSTCLIST
    Do While Not TRS.EOF
        Set List = .ListItems.Add(, , TRS.Fields(0), , 1)
        With List
            .SubItems(1) = TRS.Fields(1)
            .SubItems(2) = TRS.Fields(2)
            
        End With
        TRS.MoveNext
    Loop
    End With
End Sub

Private Sub LSTCLIST_DblClick()
    CMDOK_Click
End Sub

Private Sub LSTRESULTS_DblClick()
    CMDOK_Click
End Sub

