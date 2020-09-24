VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMDOWLOAD 
   Caption         =   "Downloading"
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5715
   Icon            =   "FRMDOWLOAD.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1335
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   240
      Top             =   240
   End
   Begin VB.CommandButton cmdProcessArray 
      Caption         =   "System Downloading "
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Top             =   240
      Width           =   2475
   End
   Begin MSComctlLib.ProgressBar prgArray 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "FRMDOWLOAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdProcessArray_Click()
    Dim Progress(500) As Integer
    Dim counter As Integer
    'We don't need to reset the progress bar because
    'its value is set in the loop
    
    'Loop through the array
    For counter = LBound(Progress) To UBound(Progress)
        Progress(counter) = counter
        'Update the value of the progress bar
        prgArray.Value = counter
    Next counter
End Sub

Private Sub Form_Load()
    With prgArray
        .Min = 0
        .Max = 500
        .Value = 0
    End With
End Sub

Private Sub Timer1_Timer()

Call cmdProcessArray_Click
Unload Me
'FRMSCREEN2.Show
End Sub
