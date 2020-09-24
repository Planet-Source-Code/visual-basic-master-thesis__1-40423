VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form FRMDAILYR 
   Caption         =   "Daily Sales Report"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   Icon            =   "FRMDAILYR.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      Begin VB.CommandButton Command1 
         Caption         =   "Print Preview"
         Height          =   375
         Left            =   6480
         TabIndex        =   1
         Top             =   2880
         Width           =   1095
      End
      Begin MSACAL.Calendar Calendar2 
         Height          =   2295
         Left            =   3960
         TabIndex        =   2
         Top             =   600
         Width           =   3735
         _Version        =   524288
         _ExtentX        =   6588
         _ExtentY        =   4048
         _StockProps     =   1
         BackColor       =   -2147483639
         Year            =   2002
         Month           =   9
         Day             =   24
         DayLength       =   0
         MonthLength     =   0
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   0
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   0   'False
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   0   'False
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   2295
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3855
         _Version        =   524288
         _ExtentX        =   6800
         _ExtentY        =   4048
         _StockProps     =   1
         BackColor       =   -2147483639
         Year            =   2002
         Month           =   9
         Day             =   24
         DayLength       =   0
         MonthLength     =   0
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   0
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   0   'False
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   0   'False
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Begin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
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
         Left            =   3960
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox Crwrpt 
      Height          =   480
      Left            =   2760
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   3000
      Width           =   1200
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   9
      Height          =   3615
      Left            =   120
      Top             =   120
      Width           =   8055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   7
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "FRMDAILYR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
With Crwrpt
        
        .DataFiles(0) = App.Path & "\Database\MasterDB.mdb"
        .ReportFileName = App.Path & "\Reports\Collection.rpt"
        .SelectionFormula = "{Collection.date} In Date (" & Format$(Calendar1.Value, "yyyy,mm,dd") & ") To Date (" & Format$(Calendar2.Value, "yyyy,mm,dd") & ")"
        .WindowBorderStyle = crptFixedSingle
        .WindowState = crptMaximized
        .WindowTitle = "Daily Collection Report"
        .Action = 1
        Screen.MousePointer = vbDefault
End With
End Sub
