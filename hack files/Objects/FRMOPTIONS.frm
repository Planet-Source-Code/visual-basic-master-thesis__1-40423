VERSION 5.00
Begin VB.Form FRMOPTIONS 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "FRMOPTIONS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   420
      Left            =   4545
      TabIndex        =   4
      Top             =   2355
      Width           =   1125
   End
   Begin VB.CommandButton CMDCHANGE 
      Caption         =   "&Change Password"
      Default         =   -1  'True
      Height          =   435
      Left            =   2460
      TabIndex        =   3
      Top             =   2340
      Width           =   1980
   End
   Begin VB.TextBox TXTNEW 
      Height          =   400
      IMEMode         =   3  'DISABLE
      Left            =   2500
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1125
      Width           =   3180
   End
   Begin VB.TextBox TXTCONFIRM 
      Height          =   400
      IMEMode         =   3  'DISABLE
      Left            =   2500
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1680
      Width           =   3180
   End
   Begin VB.TextBox TXTOLD 
      Height          =   400
      IMEMode         =   3  'DISABLE
      Left            =   2500
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   585
      Width           =   3180
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1920
      Picture         =   "FRMOPTIONS.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Change System Password"
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
      Left            =   2520
      TabIndex        =   8
      Top             =   195
      Width           =   2190
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Confirm Password"
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
      Left            =   495
      TabIndex        =   7
      Top             =   1785
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "New Password"
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
      Left            =   495
      TabIndex        =   6
      Top             =   1215
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Old Password"
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
      Left            =   495
      TabIndex        =   5
      Top             =   690
      Width           =   1170
   End
End
Attribute VB_Name = "FRMOPTIONS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDCHANGE_Click()
    If TXTNEW.Text = "" And TXTCONFIRM.Text = "" Then
        MsgBox "Please enter password to changed", vbInformation, "Confirm"
        Exit Sub
    End If
    If TXTNEW.Text <> TXTCONFIRM.Text Then
        MsgBox "Confirm password does not match", vbInformation, "Confirm"
        With TXTNEW
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        With TXTCONFIRM
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        TXTNEW.SetFocus
    End If
    Call ChangePassword(TXTOLD.Text, TXTNEW.Text)
    Command2_Click
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

