VERSION 5.00
Begin VB.Form FRMMAINMENU 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   Icon            =   "FRMAINMENU.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   4545
   Begin VB.CommandButton CMDCUSTOMER 
      Caption         =   "&Customer"
      Height          =   735
      Left            =   2520
      MouseIcon       =   "FRMAINMENU.frx":27A2
      MousePointer    =   99  'Custom
      Picture         =   "FRMAINMENU.frx":2AAC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton CMDSTOCKS 
      Caption         =   "&Stocks"
      Height          =   855
      Left            =   2520
      MouseIcon       =   "FRMAINMENU.frx":38EE
      MousePointer    =   99  'Custom
      Picture         =   "FRMAINMENU.frx":3BF8
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton CMDSUPPLIER 
      Caption         =   "&Supplier"
      Height          =   855
      Left            =   2520
      MouseIcon       =   "FRMAINMENU.frx":403A
      MousePointer    =   99  'Custom
      Picture         =   "FRMAINMENU.frx":4344
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton CMDORDERCASH 
      Caption         =   "&Order in Cash"
      Height          =   735
      Left            =   600
      MouseIcon       =   "FRMAINMENU.frx":5186
      MousePointer    =   99  'Custom
      Picture         =   "FRMAINMENU.frx":5490
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton CMDORDERCREDIT 
      Caption         =   "&Credit"
      Height          =   855
      Left            =   600
      MouseIcon       =   "FRMAINMENU.frx":5D5A
      MousePointer    =   99  'Custom
      Picture         =   "FRMAINMENU.frx":6064
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton CMDPAYMENT 
      Caption         =   "&Payment"
      Height          =   855
      Left            =   600
      MouseIcon       =   "FRMAINMENU.frx":64A6
      MousePointer    =   99  'Custom
      Picture         =   "FRMAINMENU.frx":67B0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton CMDDELIVERY 
      Caption         =   "&Delivery"
      Height          =   855
      Left            =   600
      MouseIcon       =   "FRMAINMENU.frx":6ABA
      MousePointer    =   99  'Custom
      Picture         =   "FRMAINMENU.frx":6DC4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton CMDPO 
      Caption         =   "&Purchase Order"
      Height          =   855
      Left            =   600
      MouseIcon       =   "FRMAINMENU.frx":7206
      MousePointer    =   99  'Custom
      Picture         =   "FRMAINMENU.frx":7510
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales and Inventory System"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   11
      Top             =   360
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   480
      Picture         =   "FRMAINMENU.frx":81DA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   360
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   2400
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   6015
      Left            =   120
      Picture         =   "FRMAINMENU.frx":A97C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "FRMMAINMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CMDCUSTOMER_Click()
FRMCUSTOMERS.Show
End Sub

Private Sub CMDDELIVERY_Click()
FRMDELIVERY.Show
End Sub

Private Sub CMDORDERCASH_Click()
FRMOCASH.Show
End Sub

Private Sub CMDORDERCREDIT_Click()
FRMOCREDIT.Show
End Sub

Private Sub CMDPAYMENT_Click()
FRMPAYMENTS.Show
End Sub

Private Sub CMDPO_Click()
FRMPO.Show
End Sub

Private Sub CMDSTOCKS_Click()
FRMSTOCKS.Show
End Sub

Private Sub CMDSUPPLIER_Click()
FRMSUPPLIERS.Show
End Sub



