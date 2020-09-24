VERSION 5.00
Begin VB.Form FRMINFO 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Information..."
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "FRMINFO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      Height          =   615
      Left            =   240
      Picture         =   "FRMINFO.frx":27A2
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      ItemData        =   "FRMINFO.frx":4F44
      Left            =   240
      List            =   "FRMINFO.frx":4F54
      TabIndex        =   1
      Top             =   2280
      Width           =   6135
   End
   Begin VB.CommandButton CMDCLOSE 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright: 2000-2001 Turquoisesilver Corporation"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This Product is Licensed to:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number:1111111111"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Turquoisesilver Diamond Lil Mfg"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Turquoisesilver"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3255
      Left            =   120
      Picture         =   "FRMINFO.frx":505F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "FRMINFO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub Form_activate()

'Me.Line (100, 3050)-(Me.ScaleWidth - 100, 3050), QBColor(15)
'Me.Line (100, 3020)-(Me.ScaleWidth - 100, 3020), QBColor(8)
'End Sub

Private Sub CMDCLOSE_Click()
Unload Me
End Sub

