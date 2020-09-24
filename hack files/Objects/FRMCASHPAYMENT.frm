VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMCASHPAYMENT 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payment Screen"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "FRMCASHPAYMENT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TXTPAY 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   600
      TabIndex        =   15
      Top             =   4920
      Width           =   2700
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Customer Payment ]"
      Height          =   1335
      Left            =   360
      TabIndex        =   8
      Top             =   4200
      Width           =   5175
      Begin VB.ComboBox CBOTYPE 
         Height          =   315
         ItemData        =   "FRMCASHPAYMENT.frx":000C
         Left            =   3240
         List            =   "FRMCASHPAYMENT.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Enter Payment :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Type of Payment"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3270
         TabIndex        =   10
         Top             =   390
         Width           =   1200
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5790
      Top             =   2190
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
            Picture         =   "FRMCASHPAYMENT.frx":001D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CMDCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   540
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1380
   End
   Begin MSComctlLib.ListView LSTORD 
      Height          =   2295
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item Number"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description "
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Quantity"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Amount "
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton CMDPAY 
      Caption         =   "&Pay"
      Default         =   -1  'True
      Height          =   555
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4905
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   360
      ScaleHeight     =   2835
      ScaleWidth      =   9435
      TabIndex        =   3
      Top             =   1200
      Width           =   9495
      Begin VB.Label LBLMESS 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Prime Asia And Jewelry Shop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   "00-88889-00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Tin #"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Label LBLTIME 
      Caption         =   "12:33:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label LBLMESS 
      BackColor       =   &H8000000A&
      Caption         =   "Date Ordered : 12/12/1212"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   1
      Left            =   6840
      TabIndex        =   5
      Top             =   840
      Width           =   3630
   End
End
Attribute VB_Name = "FRMCASHPAYMENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDCANCEL_Click()
FRMOCASH.LSTORDERS.ListItems.Clear
FRMOCASH.TXTTOTAL.Text = "0.00"
FRMOCASH.LBLITEM.Caption = "0"
    Unload Me
End Sub



Private Sub CMDPAY_Click()
Dim List As ListItem
Dim message As String

If Val(TXTPAY.Text) < OrdAmount Then
        MsgBox " Payment not enough ", vbInformation, "Message"
        CMDPAY.Caption = "&Pay"
Exit Sub
ElseIf Val(TXTPAY.Text) >= OrdAmount Then
   Set List = LSTORD.ListItems.Add(, , "Amount Tendered:", , 0)
   With List
   .SubItems(3) = Format(Val(TXTPAY.Text), "###,###,##0.00")
   End With
   Set List = LSTORD.ListItems.Add(, , "Change:", , 0)
   With List
   
   .SubItems(3) = Format(Str(Val(TXTPAY.Text) - OrdAmount), "###,###,##0.00")
   End With
End If
Dim Ordnumber As String
Dim Query As String
Dim x As Long
Ordnumber = LBLMESS(0).Tag
    
    'Add all ordered stocks to orders
    For x = 1 To FRMOCASH.LSTORDERS.ListItems.Count
        Query = "INSERT INTO orders (ordno,custno,stockno,invno,amount,quantity) VALUES " _
             & "('" & Ordnumber & "','Cash','" & FRMOCASH.LSTORDERS.ListItems(x).Text & "','DONT'," _
             & FRMOCASH.LSTORDERS.ListItems(x).SubItems(3) & "," & FRMOCASH.LSTORDERS.ListItems(x).SubItems(2) & ");"
        DBMain.Execute Query
    Next x
    'Add records to payments
    Query = "INSERT INTO sales (ordno,amount,[date],paytype,type) VALUES ('" & Ordnumber & "'," & OrdAmount & ",#" & Format(Now, "mm/dd/yy") & "#,'" & CBOTYPE.List(CBOTYPE.ListIndex) & "','Cash');"
    DBMain.Execute Query
    OcashStat = "OK"
    MsgBox "Payment Record has been successfully Added", vbExclamation, "Payments"
    TXTPAY.Text = ""
    CMDPAY.Enabled = False
    FRMOCASH.LSTORDERS.ListItems.Clear
    FRMOCASH.LBLITEM.Caption = 0
    FRMOCASH.TXTTOTAL.Text = "0.00"
    FRMOCASH.LSTORDERS.ListItems.Clear
    FRMOCASH.TXTTOTAL.Text = "0.00"
    FRMOCASH.LBLITEM.Caption = "0"
End Sub

Private Sub Form_Load()
    LBLTIME.Caption = Format(Now, " HH:MM:SS AMPM")
CBOTYPE.ListIndex = 0
End Sub
Private Sub TXTPAY_Change()
    If TXTPAY.Text <> "" Then
        CMDPAY.Enabled = True
    Else
        CMDPAY.Enabled = False
    End If
End Sub

' this procedure will accept only Numbers and Period
Private Sub TXTPAY_KeyPress(KeyAscii As Integer)

If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 8) Then KeyAscii = 0
Exit Sub
If KeyAscii = 13 Then
CMDPAY.Value = True
End If
End Sub
