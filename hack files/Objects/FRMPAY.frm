VERSION 5.00
Begin VB.Form FRMPAY 
   BackColor       =   &H8000000A&
   Caption         =   "Enter Payment"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2505
   ScaleWidth      =   5520
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CMDCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   450
      Left            =   3960
      TabIndex        =   4
      Top             =   1605
      Width           =   1400
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   450
      Left            =   3960
      TabIndex        =   3
      Top             =   1065
      Width           =   1400
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "[ Payment Type ]"
      ForeColor       =   &H00C00000&
      Height          =   1440
      Left            =   210
      TabIndex        =   2
      Top             =   960
      Width           =   3660
      Begin VB.OptionButton OPTTYPE 
         BackColor       =   &H8000000A&
         Caption         =   "Check"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   200
         TabIndex        =   6
         Tag             =   "Check"
         Top             =   990
         Width           =   1875
      End
      Begin VB.OptionButton OPTTYPE 
         BackColor       =   &H8000000A&
         Caption         =   "Cash Payment"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   200
         TabIndex        =   5
         Tag             =   "Cash"
         Top             =   540
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.TextBox TXTAMOUNT 
      Alignment       =   1  'Right Justify
      Height          =   400
      Left            =   1800
      TabIndex        =   1
      Top             =   330
      Width           =   3525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Enter Payment :"
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
      Left            =   390
      TabIndex        =   0
      Top             =   435
      Width           =   1365
   End
End
Attribute VB_Name = "FRMPAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PType As String

Private Sub CMDCANCEL_Click()
    Unload Me
End Sub

Private Sub CMDOK_Click()
    Dim TQR As DAO.QueryDef
    Dim TRS As DAO.Recordset
    Dim Query As String
    RecordPayment Me.Tag, TXTAMOUNT.Text, PType
    MsgBox "Payment Record has been Successfully Added", vbInformation, "Payment"
    Query = "SELECT ordno,amount,date,type FROM payments WHERE custno='" & Me.Tag & "';"
    Set TQR = DBMain.CreateQueryDef("", Query)
    Set TRS = TQR.OpenRecordset()
    DisplayRecord FRMPAYMENTS.LSTPAYMENTS, TRS, 3, 1
    FRMPAYMENTS.TXTACCBAL = Format(AccBalance(Me.Tag), "###,###,##0.00")
    Set TQR = Nothing
    Set TRS = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        CMDCANCEL_Click
    End If
End Sub


Private Sub Form_Load()
    PType = "Cash"
End Sub

Private Sub OPTTYPE_Click(Index As Integer)
    PType = OPTTYPE(Index).Tag
End Sub

Private Sub TXTAMOUNT_Change()
    If TXTAMOUNT.Text = "" Then
        CMDOK.Enabled = False
    Else
        CMDOK.Enabled = True
    End If
End Sub

Private Sub TXTAMOUNT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 46 Then KeyAscii = 0
End Sub
