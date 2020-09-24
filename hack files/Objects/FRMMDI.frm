VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form FRMMDI 
   BackColor       =   &H80000007&
   Caption         =   "Sales And Inventory System"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ForeColor       =   &H00FFFF00&
   Icon            =   "FRMMDI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3720
      Top             =   8160
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      Height          =   6255
      Left            =   480
      ScaleHeight     =   6195
      ScaleWidth      =   10515
      TabIndex        =   0
      Top             =   1800
      Width           =   10575
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
         Height          =   6015
         Left            =   120
         TabIndex        =   2
         Top             =   105
         Width           =   10335
         Movie           =   ""
         Src             =   ""
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   -1  'True
         Base            =   ""
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   "000000"
      End
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000F&
      BorderWidth     =   3
      X1              =   11040
      X2              =   11040
      Y1              =   1680
      Y2              =   1320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000F&
      BorderWidth     =   3
      X1              =   10800
      X2              =   11040
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000F&
      BorderWidth     =   3
      X1              =   240
      X2              =   11040
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000F&
      BorderWidth     =   3
      X1              =   240
      X2              =   240
      Y1              =   1320
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000F&
      BorderWidth     =   3
      X1              =   240
      X2              =   480
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sales and Inventory System"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   10335
   End
End
Attribute VB_Name = "FRMMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function PlayFlashMovie(Filename As String)
    With Flash1
        .Movie = Filename
        .Play
    End With
End Function


Public Sub Form_Activate()
On Error Resume Next
DrawBorder 100, Me
PlayFlashMovie ("C:\Ted Antonio File\Microsoft Visual Basic 5.0\DATABASE_PROGRAM\SISystem\Objects\MILINTRO.SWF")

End Sub

Private Sub Timer2_Timer()
Unload Me
End Sub
Dim x(100), Y(100), pace(100), size(100) As Integer
Private Sub cmdQuit_Click()
'Unload Me is actually a better way to end your
'program than End, since End may leave parts of
'the program in memory.
Unload Me
End Sub
'Private Sub Form_Activate()
'Randomize
''Randomize is a statement that along with Rnd allows
''you to generate Random numbers. Randomize initializes
''this random number generator based on a value from your
''system timer.
'For i = 1 To 100
''The Int function returns the integer portion of the
''number passed to it. For example X1 = Int(99.8) will
''return a value of 99 to X1, so will Int(99.1).
'X1 = Int(Form1.Width * Rnd)
'Y1 = Int(Form1.Height * Rnd)
'
''The idea of this pace is to generate a random speed
''as it goes through the loop. Since it goes through the
''loop so fast you may not notice the changes. If you want
''to experiment with this then comment out the pace1 = Int..
''line and uncomment the pace1 = 0. Try changing the value
''of zero and watch the speed change. You may prefer creating
''a variable called velocity and then set pace1=velocity, this
''way you could control the speed by setting new value for
''velocity.
'
'pace1 = Int(500 - (Int(Rnd * 499)))
''pace1 = 0
'
''This next piece assigns a random value to size which is
''passed to the Circle method in the Timer1_Timer event
''resulting in different size circles. You can increase
''the max size of the by changing the 25 but the circles
''will not be filled. You'll have to do some fiddling on
''your own to fill the circles without drawing tracks on
''background.
'size1 = 25 * Rnd
'x(i) = X1
'Y(i) = Y1
'pace(i) = pace1
'size(i) = size1
'Next
'End Sub
Private Sub Timer1_Timer()
For i = 1 To 100
Circle (x(i), Y(i)), size(i), BackColor
Y(i) = Y(i) + pace(i)
If Y(i) >= Form1.Height Then Y(i) = 0: x(i) = Int(Form1.Width * Rnd)
Circle (x(i), Y(i)), size(i)
Next
End Sub
Private Sub Form_click()
Unload Me
End Sub

