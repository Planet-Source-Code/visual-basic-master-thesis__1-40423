VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMPAYMENTS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payments"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   Icon            =   "FRMPAYMENTS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TXTDUEAMOUNT 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8640
      TabIndex        =   19
      Top             =   960
      Width           =   2505
   End
   Begin VB.CommandButton CMDCLOSE 
      Caption         =   "Close"
      Height          =   450
      Left            =   10080
      TabIndex        =   15
      Top             =   4170
      Width           =   1335
   End
   Begin MSComctlLib.ListView LSTPAYMENTS 
      Height          =   2625
      Left            =   285
      TabIndex        =   4
      Top             =   3720
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   4630
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Order Number"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Amount Paid"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date_Paid"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Payment Type"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3210
      Left            =   150
      TabIndex        =   3
      Top             =   3240
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   5662
      TabFixedHeight  =   353
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "          Payment Details          "
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "[ Customer ]"
      ForeColor       =   &H00000000&
      Height          =   2880
      Left            =   150
      TabIndex        =   1
      Top             =   105
      Width           =   11295
      Begin VB.TextBox TXTACCBAL 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8520
         TabIndex        =   17
         Top             =   1320
         Width           =   2505
      End
      Begin VB.CommandButton CMDBROWSE 
         Height          =   285
         Left            =   4350
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   360
         Width           =   285
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1995
         TabIndex        =   12
         Top             =   1800
         Width           =   3915
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2000
         TabIndex        =   7
         Top             =   840
         Width           =   2760
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2000
         TabIndex        =   6
         Top             =   1320
         Width           =   2790
      End
      Begin VB.TextBox TXTF 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2000
         TabIndex        =   5
         Top             =   2280
         Width           =   1920
      End
      Begin VB.TextBox TXTF 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2000
         TabIndex        =   2
         Top             =   360
         Width           =   2280
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Amount Due:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6495
         TabIndex        =   20
         Top             =   840
         Width           =   930
      End
      Begin VB.Label LBLM 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Running Balance :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   6480
         TabIndex        =   18
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "[ Accounts ]"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8385
         TabIndex        =   14
         Top             =   15
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         Index           =   1
         X1              =   6300
         X2              =   6300
         Y1              =   135
         Y2              =   2760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         Index           =   0
         X1              =   6285
         X2              =   6285
         Y1              =   120
         Y2              =   2760
      End
      Begin VB.Label LBLM 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Address :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   13
         Top             =   1920
         Width           =   660
      End
      Begin VB.Label LBLM 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Credit Limit :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label LBLM 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         Caption         =   "Firstname :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   10
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label LBLM 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Lastname :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   9
         Top             =   960
         Width           =   780
      End
      Begin VB.Label LBLM 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "Customer Number : "
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   8
         Top             =   450
         Width           =   1395
      End
   End
   Begin VB.CommandButton CMDPAYMENTS 
      Caption         =   "&Payment"
      Enabled         =   0   'False
      Height          =   450
      Left            =   10080
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
   Begin MSComctlLib.ImageList IMGLIST 
      Left            =   10515
      Top             =   4665
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
            Picture         =   "FRMPAYMENTS.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FRMPAYMENTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TQR As DAO.QueryDef
Dim TRS As DAO.Recordset
Dim Query As String

Private Sub CMDBROWSE_Click()
    FRMPAYCUSTOMER.Show vbModal
End Sub

Private Sub CMDCLOSE_Click()
    Unload Me

End Sub


Private Sub CMDPAYMENTS_Click()
    CustPay = GetOrders(TXTF(0).Text)
    If CustPay.ordno = "" Then
        MsgBox "Customer has no Account to pay", vbInformation, "System"
        Exit Sub
    End If
    FRMPAY.Tag = TXTF(0).Text
    FRMPAY.Caption = "Enter Payment - " & TXTF(0).Text
    FRMPAY.Show vbModal
End Sub

Private Sub Form_Load()
    SetFlatList LSTPAYMENTS, Me
End Sub

Private Sub TXTF_Change(Index As Integer)
    If Index = 0 Then
        Query = "SELECT ordno,amount,date,type FROM payments WHERE custno='" & TXTF(0).Text & "';"
        Set TQR = DBMain.CreateQueryDef("", Query)
        Set TRS = TQR.OpenRecordset()
        DisplayRecord LSTPAYMENTS, TRS, 3, 1
        TXTACCBAL.Text = Format(AccBalance(TXTF(0).Text), "###,###,##0.00")
        AmountDue TXTF(0).Text
    If TXTF(0).Text = "" Then
        CMDPAYMENTS.Enabled = False
    Else
        CMDPAYMENTS.Enabled = True
    End If
    End If
End Sub
