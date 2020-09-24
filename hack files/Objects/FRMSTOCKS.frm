VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMSTOCKS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "Item MasterFile"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "FRMSTOCKS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "[ Item Information ]"
      ForeColor       =   &H00000000&
      Height          =   6615
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   11175
      Begin VB.CommandButton CMDFIND 
         Caption         =   "&Find"
         Height          =   800
         Left            =   9600
         Picture         =   "FRMSTOCKS.frx":000C
         TabIndex        =   17
         Top             =   2625
         Width           =   1300
      End
      Begin VB.CommandButton CMDCLOSE 
         Caption         =   "Close"
         Height          =   800
         Left            =   9600
         Picture         =   "FRMSTOCKS.frx":010E
         TabIndex        =   16
         Top             =   3480
         Width           =   1300
      End
      Begin VB.CommandButton CMDSAVE 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   800
         Left            =   9600
         Picture         =   "FRMSTOCKS.frx":02AB
         TabIndex        =   15
         Top             =   1800
         Width           =   1300
      End
      Begin VB.CommandButton CMDMODIFY 
         Caption         =   "&Modify"
         Height          =   800
         Left            =   9600
         Picture         =   "FRMSTOCKS.frx":0915
         TabIndex        =   14
         Top             =   1035
         Width           =   1300
      End
      Begin VB.CommandButton CMDADD 
         Caption         =   "&Add"
         Height          =   800
         Left            =   9600
         Picture         =   "FRMSTOCKS.frx":0D57
         TabIndex        =   13
         Top             =   240
         Width           =   1300
      End
      Begin VB.TextBox TXTF 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   0
         Left            =   1695
         MaxLength       =   6
         TabIndex        =   7
         Top             =   240
         Width           =   2445
      End
      Begin VB.TextBox TXTF 
         Height          =   285
         Index           =   1
         Left            =   1725
         MaxLength       =   30
         TabIndex        =   6
         Top             =   720
         Width           =   5445
      End
      Begin VB.TextBox TXTF 
         Height          =   285
         Index           =   2
         Left            =   1710
         TabIndex        =   5
         Top             =   1200
         Width           =   2520
      End
      Begin VB.TextBox TXTF 
         Height          =   285
         Index           =   3
         Left            =   1695
         TabIndex        =   4
         Top             =   1680
         Width           =   2505
      End
      Begin VB.TextBox TXTF 
         Height          =   285
         Index           =   4
         Left            =   1680
         TabIndex        =   3
         Top             =   2160
         Width           =   2505
      End
      Begin MSComctlLib.ListView LSTSTOCKS 
         Height          =   3075
         Left            =   240
         TabIndex        =   1
         Top             =   3360
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5424
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "IMGSTOCKS"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Number"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description "
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Unit Price"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Re-Order Level"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3660
         Left            =   120
         TabIndex        =   2
         Top             =   2880
         Width           =   9060
         _ExtentX        =   15981
         _ExtentY        =   6456
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "       Items      "
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label LBLCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Item Number :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   12
         Top             =   360
         Width           =   990
      End
      Begin VB.Label LBLCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Description   :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label LBLCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Unit Price     :"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label LBLCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Quantity      :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label LBLCAP 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Re-Order level :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   1110
      End
   End
   Begin MSComctlLib.ImageList IMGSTOCKS 
      Left            =   2190
      Top             =   3930
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
            Picture         =   "FRMSTOCKS.frx":1621
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   8
      Height          =   6855
      Left            =   240
      Top             =   360
      Width           =   11415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   7
      Height          =   7095
      Left            =   120
      Top             =   240
      Width           =   11655
   End
End
Attribute VB_Name = "FRMSTOCKS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DBMain As DAO.Database
Public WSMain As DAO.Workspace
Dim MyType As String

Private Sub CMDADD_Click()
   
    CMDSAVE.Enabled = True
    'CMDCANCEL.Enabled = True
    
    CMDADD.Enabled = False
    'CMDDELETE.Enabled = False
    CMDMODIFY.Enabled = False
    CMDCLOSE.Enabled = False
    CMDFIND.Enabled = False
    TS True
    TXTF(0).Text = AutoStock
    LSTSTOCKS.Enabled = False
   ClearTextE
    
    MyType = "ADD"
End Sub



Private Sub CMDCLOSE_Click()
    Unload Me
    End Sub



Private Sub CMDFIND_Click()
    FRMSTOCKSFIND.Show vbModal
End Sub

Private Sub CMDMODIFY_Click()
    If TXTF(0).Text = "" Then
        MsgBox "Please select a record to be Modified", vbExclamation, "Message"
        Exit Sub
    End If
    CMDSAVE.Enabled = True
    'CMDCANCEL.Enabled = True
    
    CMDADD.Enabled = False
    'CMDDELETE.Enabled = False
    CMDMODIFY.Enabled = False
    CMDFIND.Enabled = False
    CMDCLOSE.Enabled = False
    
    TS True
    TXTF(1).SetFocus
    LSTSTOCKS.Enabled = False
    MyType = "EDIT"
End Sub

Private Sub CMDSAVE_Click()
    Dim List As ListItem
    Dim Query As String
    Dim Flag  As Boolean
    Dim x As Long
    Flag = False
    For x = 0 To 4
        If TXTF(x).Text = "" Then Flag = True
    Next x
    If Flag Then
        MsgBox "Please enter Complete information about the items", vbInformation, "Confirm"
        GoTo last
    End If
    If MyType = "ADD" Then
        Query = "INSERT INTO stocks (stockno,description,unitprice,quantity,reorder) VALUES ('" & TXTF(0).Text & "','" & TXTF(1).Text & "'," & TXTF(2).Text & "," & TXTF(3).Text & "," & TXTF(4).Text & ");"
        DBMain.Execute Query
        Set List = LSTSTOCKS.ListItems.Add(, , TXTF(0).Text, , 1)
        With List
            .SubItems(1) = TXTF(1).Text
            .SubItems(2) = Format(TXTF(2).Text, "###,###,##0.00")
            .SubItems(3) = TXTF(3).Text
            .SubItems(4) = TXTF(4).Text
        End With
        'CMDCANCEL_Click
    ElseIf MyType = "EDIT" Then
        Query = "UPDATE stocks SET description ='" & TXTF(1).Text & "', unitprice =" & Val(Str(TXTF(2).Text)) & ", quantity = " & Val(TXTF(3).Text) & ", reorder = " & TXTF(4).Text & " WHERE stockno ='" & TXTF(0).Text & "';"
        DBMain.Execute Query
        TS False
        Set List = LSTSTOCKS.FindItem(TXTF(0).Text, , , lvwPartial)
        With List
            .SubItems(1) = TXTF(1).Text
            .SubItems(2) = Format(TXTF(2).Text, "###,###,##0.00")
            .SubItems(3) = TXTF(3).Text
            .SubItems(4) = TXTF(4).Text
        End With
        CMDADD.Enabled = True
        'CMDDELETE.Enabled = True
        CMDMODIFY.Enabled = True
        CMDFIND.Enabled = True
        CMDCLOSE.Enabled = True
        CMDSAVE.Enabled = False
        'CMDCANCEL.Enabled = True
        LSTSTOCKS.Enabled = True
    End If
    Exit Sub
last:
    TXTF(1).SetFocus
    

End Sub


Private Sub TS(ByVal Status As Boolean)
    Dim x As Long
    For x = 1 To 4
        TXTF(x).Enabled = Status
    Next x
End Sub

Private Sub Clear()
    Dim i As Long
    For i = 0 To 4
        TXTF(i).Text = ""
    Next i
End Sub

Private Sub Form_Load()
  Set WSMain = DBEngine.Workspaces(0)
    ' Open Main Database
    Set DBMain = WSMain.OpenDatabase(App.Path + "\Database\MasterDB.mdb")
TS False
FillStocks
End Sub

Private Sub LSTSTOCKS_ItemClick(ByVal Item As MSComctlLib.ListItem)
    TXTF(0).Text = Item.Text
    TXTF(1).Text = Item.SubItems(1)
    TXTF(2).Text = Format(Item.SubItems(2), "###,###,##0.00")
    TXTF(3).Text = Item.SubItems(3)
    TXTF(4).Text = Item.SubItems(4)
End Sub

