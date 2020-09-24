VERSION 5.00
Begin VB.Form FRMDLG 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quantity"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CMDCANCEL 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   3840
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   405
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox TXTVALUE 
      Alignment       =   1  'Right Justify
      Height          =   400
      Left            =   1200
      TabIndex        =   1
      Top             =   645
      Width           =   2445
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Enter Number of Quantity"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   195
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Quantity :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   735
      Width           =   840
   End
End
Attribute VB_Name = "FRMDLG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDCANCEL_Click()
    Qty = 0
Unload Me
End Sub

Private Sub CMDOK_Click()
    Qty = Val(TXTVALUE.Text)
    FRMOCASH.LBLITEM.Caption = Qty + Val(FRMOCASH.LBLITEM.Caption)
       Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
       Unload Me
    End If
End Sub

