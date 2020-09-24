VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMSUPPLIERBROWSE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suppliers"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "FRMSUPPLIERBROWSE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7380
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CMDSELECT 
      Caption         =   "Select"
      Height          =   390
      Left            =   4890
      TabIndex        =   3
      Top             =   3285
      Width           =   1155
   End
   Begin VB.CommandButton CMDCANCEL 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   6090
      TabIndex        =   2
      Top             =   3300
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "[ Available Supplier(s) ]"
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   135
      Width           =   7230
      Begin MSComctlLib.ListView ListView1 
         Height          =   2550
         Left            =   105
         TabIndex        =   1
         Top             =   345
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   4498
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "FRMSUPPLIERBROWSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

