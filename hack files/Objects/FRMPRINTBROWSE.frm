VERSION 5.00
Begin VB.Form FRMPRINTCUSTOMER 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Information"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "FRMPRINTBROWSE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6270
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CMDVIEW 
      Caption         =   "..."
      Height          =   420
      Left            =   5535
      MouseIcon       =   "FRMPRINTBROWSE.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   255
      Width           =   465
   End
   Begin VB.TextBox TXTF 
      Enabled         =   0   'False
      Height          =   400
      Index           =   3
      Left            =   2000
      TabIndex        =   9
      Top             =   1785
      Width           =   4000
   End
   Begin VB.TextBox TXTF 
      Enabled         =   0   'False
      Height          =   400
      Index           =   2
      Left            =   2000
      TabIndex        =   8
      Top             =   1290
      Width           =   4000
   End
   Begin VB.TextBox TXTF 
      Enabled         =   0   'False
      Height          =   400
      Index           =   1
      Left            =   2000
      TabIndex        =   7
      Top             =   780
      Width           =   4000
   End
   Begin VB.TextBox TXTF 
      Enabled         =   0   'False
      Height          =   400
      Index           =   0
      Left            =   2000
      TabIndex        =   6
      Top             =   255
      Width           =   3495
   End
   Begin VB.CommandButton CMDCANCEL 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   405
      Left            =   4905
      TabIndex        =   1
      Top             =   2400
      Width           =   1140
   End
   Begin VB.CommandButton CMDPREVIEW 
      Caption         =   "    &Preview"
      Enabled         =   0   'False
      Height          =   405
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1140
   End
   Begin VB.Label LBLMES 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Address"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   200
      TabIndex        =   5
      Top             =   1905
      Width           =   570
   End
   Begin VB.Label LBLMES 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Last Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   200
      TabIndex        =   4
      Top             =   1395
      Width           =   765
   End
   Begin VB.Label LBLMES 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "First Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   195
      TabIndex        =   3
      Top             =   840
      Width           =   750
   End
   Begin VB.Label LBLMES 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "Enter Customer Number"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   200
      TabIndex        =   2
      Top             =   360
      Width           =   1680
   End
End
Attribute VB_Name = "FRMPRINTCUSTOMER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDCANCEL_Click()
    Unload Me
End Sub

Private Sub CMDPREVIEW_Click()
    With MDIMAIN.CRWReport
        Screen.MousePointer = vbHourglass
        .WindowState = crptMaximized
        .SelectionFormula = "{Statement.Custno} = '" & TXTF(0).Text & "'"
        .WindowBorderStyle = crptFixedSingle
        .DataFiles(0) = App.Path & "\Database\Masterdb.mdb"
        .WindowTitle = "Customer Statement of Account Report"
        .ReportFileName = App.Path & "\Reports\Statement.rpt"
        .Action = 1
        Screen.MousePointer = vbDefault
    End With
End Sub

Private Sub CMDVIEW_Click()
    FRMVIEWCUST.Show vbModal
End Sub

Private Sub TXTF_Change(Index As Integer)
    If Index = 0 Then
        If TXTF(0).Text <> "" Then
            CMDPREVIEW.Enabled = True
        Else
            CMDPREVIEW.Enabled = False
        End If
    End If
End Sub
'**************************************
' Name: crystal report date range
' Description:Print a crystal report in
'     a certain date range
' By: James Ramsay
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/xq/ASP/txtCode
'     Id.14757/lngWId.1/qx/vb/scripts/ShowCode
'     .htm'for details.'**************************************

'report.ReportFileName = gvPath & "\cheques.rpt"
'report.CopiesToPrinter = InputBox("How many copies would you like to print")
'report.SelectionFormula = "{cheques.date} In Date (" & Format$(Startdatetextbox.Value, "yyyy,mm,dd") & ") To Date (" & Format$(enddatetextbox.Value, "yyyy,mm,dd") & ")"
'report.ReportTitle = "Report between" & " " & Format$(Startdatetextbox.Value, "long date") & " " & "and" & " " & Format(enddatetextbox.Value, "long date")
'report.Action = 1

