VERSION 5.00
Begin VB.Form FRMORD_NO 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "O.R Number"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMORD_NO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox TXTOR 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Enter O.R Number"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "FRMORD_NO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub Command1_Click()
Unload Me
Unload FRMDAMAGED

End Sub

Private Sub Command2_Click()

If TXTOR.Text = "" Then GoTo errhandler:
ord = TXTOR.Text
FRMDAMAGED.Caption = FRMDAMAGED.Caption & ":" & " Order number:" & ord
Unload Me
Exit Sub
errhandler:
MsgBox "Make sure you Enter the Order Number,Don't leave it blank", vbOKOnly, "error"
End Sub



Private Sub TXTOR_KeyPress(KeyAscii As Integer)
Integers1 KeyAscii
End Sub
