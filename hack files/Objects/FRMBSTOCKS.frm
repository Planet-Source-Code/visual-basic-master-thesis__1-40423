VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMBSTOCKS 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse Stocks"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   Icon            =   "FRMBSTOCKS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   9615
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CMDCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   450
      Left            =   8280
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton CMDSELECT 
      Caption         =   "Select"
      Height          =   450
      Left            =   6960
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame FRASTOCKS 
      BackColor       =   &H8000000A&
      Caption         =   "[ Stocks ]"
      ForeColor       =   &H00000000&
      Height          =   3990
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   9420
      Begin MSComctlLib.ImageList IMGLIST 
         Left            =   5700
         Top             =   2340
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
               Picture         =   "FRMBSTOCKS.frx":000C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView LSTSTOCKS 
         Height          =   3540
         Left            =   120
         TabIndex        =   3
         Top             =   315
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   6244
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Number"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   11642
         EndProperty
      End
   End
End
Attribute VB_Name = "FRMBSTOCKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim List As ListItem

Private Sub CMDCANCEL_Click()
    Unload Me
End Sub

Private Sub CMDSELECT_Click()
    Dim Qty As Long
    Qty = GetQuantity()
    If Qty = 0 Then Exit Sub
    Set List = FRMPO.LSTPO.FindItem(LSTSTOCKS.SelectedItem.Text, , , lvwPartial)
    If List Is Nothing Then
        Set List = FRMPO.LSTPO.ListItems.Add(, , LSTSTOCKS.SelectedItem.Text, , 1)
        With List
          .SubItems(1) = LSTSTOCKS.SelectedItem.SubItems(1)
          .SubItems(2) = Qty
        End With
    Else
        List.SubItems(2) = Qty
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Set TQR = DBMain.QueryDefs("BStocks")
    Set TRS = TQR.OpenRecordset()
    SetFlatList LSTSTOCKS, Me
    DisplayRecord LSTSTOCKS, TRS, 1, 1
End Sub

Private Sub LSTSTOCKS_DblClick()
    CMDSELECT_Click
End Sub
