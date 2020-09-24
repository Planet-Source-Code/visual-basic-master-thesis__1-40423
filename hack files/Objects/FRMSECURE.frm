VERSION 5.00
Begin VB.Form FRMSECURE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Clerk Name"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "FRMSECURE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTPASSWORD 
      Height          =   400
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1320
      Width           =   3045
   End
   Begin VB.Timer TMRSCROLL 
      Interval        =   100
      Left            =   2880
      Top             =   2520
   End
   Begin VB.Label LBLLOGIN 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      ToolTipText     =   "Click here to login"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label LBLOK 
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      ToolTipText     =   "Click here to shutdown"
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   1080
      Picture         =   "FRMSECURE.frx":000C
      Stretch         =   -1  'True
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Clerk name:"
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
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1530
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1335
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   5655
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   120
      Picture         =   "FRMSECURE.frx":08D6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6135
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   480
      X2              =   3480
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   480
      X2              =   480
      Y1              =   1920
      Y2              =   2520
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   5760
      X2              =   6000
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   6000
      X2              =   6000
      Y1              =   1320
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   5760
      X2              =   6000
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "FRMSECURE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetSystemMenu Lib "user32" _
     (ByVal hwnd As Long, _
     ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" _
     (ByVal hMenu As Long) _
     As Long
Private Declare Function RemoveMenu Lib "user32" _
     (ByVal hMenu As Long, ByVal _
     nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" _
     (ByVal hwnd As Long) As Long

Private Const MF_BYPOSITION = &H400&


Private Sub LBLOK_Click()

If TXTPASSWORD.Text = "" Then
MsgBox "Please Re Enter Clerk name:", vbOKOnly + vbInformation, "Message"
Exit Sub
Else
MDIMAIN.StatusBar1.Panels(3).Text = "Clerk:" & TXTPASSWORD.Text
Unload Me
End If
End Sub

    
Private Sub LBLLOGIN_Click()

Unload Me
Unload MDIMAIN
End Sub

'creat a line
'Private Sub Form_activate()
    'Me.Line (100, 1800)-(Me.ScaleWidth - 100, 1800), QBColor(15)
    'Me.Line (100, 1770)-(Me.ScaleWidth - 100, 1770), QBColor(8)
 

 'End Sub
'Private Sub TMRSCROLL_Timer()
'Allow the text to move horizontaly.
'Call Scroll(" Insert your password for identification. System Password  ")
'End Sub

'Private Sub Form_Load()
'Static i As Integer
'TextUpper TXTPASSWORD
'i = 0
'DisableX
'End Sub
Public Sub DisableX()
     Dim hMenu As Long
     Dim nCount As Long
     hMenu = GetSystemMenu(Me.hwnd, 0)
     nCount = GetMenuItemCount(hMenu)

     'Get rid of the Close menu and its separator
     Call RemoveMenu(hMenu, nCount - 1, MF_BYPOSITION)
     Call RemoveMenu(hMenu, nCount - 2, MF_BYPOSITION)

     'Make sure the screen updates
     'our change
     DrawMenuBar Me.hwnd
End Sub


Private Sub LBLSHUTDOWN_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    LBLOK.BackColor = &HFF&       'change the color of the background when mouse hovers over
    LBLOK.ForeColor = &H0&        'change the color of the text when mouse hovers over
    
    LBLLOGIN.BackColor = &H0&
    LBLLOGIN.ForeColor = &HFF&
    
End Sub

Private Sub LBLLOGIN_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    LBLLOGIN.BackColor = &HFF& 'change the color of the background when mouse hovers over
    LBLLOGIN.ForeColor = &H0& 'change the color of the text when mouse hovers over
    
    LBLOK.BackColor = &H0&
    LBLOK.ForeColor = &HFF&
    
   End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    LBLLOGIN.BackColor = &H0&        'change the color of the background when mouse hovers over
    LBLLOGIN.ForeColor = &HFF& 'change the color of the text when mouse hovers over
    
    LBLOK.BackColor = &H0&
    LBLOK.ForeColor = &HFF&
    
   End Sub
