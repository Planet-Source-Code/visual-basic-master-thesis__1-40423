VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMDELIVERY 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "Deliveries"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "FRMDELIVERY.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "[ Delivery Information ]"
      ForeColor       =   &H00C00000&
      Height          =   2400
      Left            =   435
      TabIndex        =   4
      Top             =   960
      Width           =   10770
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox TXTF 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   2500
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   8
         Top             =   840
         Width           =   3120
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2400
         TabIndex        =   7
         Top             =   1320
         Width           =   3120
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   8100
         TabIndex        =   6
         Top             =   1080
         Width           =   2500
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   8100
         TabIndex        =   5
         Top             =   555
         Width           =   2500
      End
      Begin VB.Label LBLMES 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Supplier Number :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   1260
      End
      Begin VB.Label LBLMES 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Supplier Name :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label LBLMES 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Telephone # :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label LBLMES 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Date Delivered :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   6795
         TabIndex        =   11
         Top             =   1050
         Width           =   1155
      End
      Begin VB.Label LBLMES 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Delivery Number :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   6675
         TabIndex        =   10
         Top             =   540
         Width           =   1260
      End
   End
   Begin VB.Frame FRASTOCKS 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   480
      TabIndex        =   1
      Top             =   3840
      Width           =   10860
      Begin VB.CommandButton CMDADD 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   540
         Left            =   9240
         Picture         =   "FRMDELIVERY.frx":000C
         TabIndex        =   18
         Top             =   240
         Width           =   1300
      End
      Begin VB.CommandButton CMDSAVE 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   540
         Left            =   9240
         Picture         =   "FRMDELIVERY.frx":044E
         TabIndex        =   17
         Top             =   2640
         Width           =   1300
      End
      Begin VB.CommandButton CMDREMOVE 
         Caption         =   "Remove"
         Enabled         =   0   'False
         Height          =   540
         Left            =   9255
         Picture         =   "FRMDELIVERY.frx":0AB8
         TabIndex        =   16
         Top             =   840
         Width           =   1300
      End
      Begin MSComctlLib.ImageList IMGLIST 
         Left            =   5865
         Top             =   945
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
               Picture         =   "FRMDELIVERY.frx":0EFA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView LSTDEL 
         Height          =   3045
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   5371
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
            Text            =   "Item Number"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   9260
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Quantity"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip TABS 
      Height          =   3675
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   6482
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "          Items    "
            Object.Tag             =   "stocks"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   6615
      Left            =   240
      ScaleHeight     =   6555
      ScaleWidth      =   11235
      TabIndex        =   3
      Top             =   600
      Width           =   11295
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   8
      Height          =   6855
      Left            =   120
      Top             =   480
      Width           =   11535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   8
      Height          =   7095
      Left            =   0
      Top             =   360
      Width           =   11775
   End
End
Attribute VB_Name = "FRMDELIVERY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDADD_Click()
    FRMDSTOCKS.Show vbModal
End Sub

Private Sub CMDBROWSE_Click()
    FRMPOBROWSE.Show vbModal
End Sub

Private Sub CMDREMOVE_Click()
    If LSTDEL.ListItems.Count = 0 Then Exit Sub
    LSTDEL.ListItems.Remove LSTDEL.SelectedItem.Index
End Sub

Private Sub CMDSAVE_Click()
    Dim Query As String
    Dim x As Long
    If LSTDEL.ListItems.Count = 0 Then
        MsgBox "There is no Item to be Recorded", vbInformation, "Delivery"
        Exit Sub
    End If
    With LSTDEL
        For x = 1 To .ListItems.Count
            RecordDelivery .ListItems(x).Text, TXTF(1).Text, .ListItems(x).SubItems(2), TXTF(4).Text, TXTF(5).Text
        Next x
    End With
    MsgBox "Delivery has been Successfully recorded ", vbInformation, "Delivery"
    Unload Me
End Sub

Private Sub Command1_Click()
FRMBROWSE.Show
End Sub


Private Sub TXTF_Change(Index As Integer)
    If Index = 1 Then
        If TXTF(1).Text <> "" Then
            TXTF(4).Text = Format(Now, "mm/dd/yyyy")
            CMDADD.Enabled = True
            CMDREMOVE.Enabled = True
            CMDSAVE.Enabled = True
        End If
    End If
End Sub


