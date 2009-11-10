VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help"
   ClientHeight    =   7035
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   9720
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   4080
      TabIndex        =   0
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9810
   End
   Begin VB.Menu mnuCpoy 
      Caption         =   "Actions"
      Begin VB.Menu mnuCopyHelpClip 
         Caption         =   "&Copy help to clipboard"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuClose 
         Caption         =   "C&lose"
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Height = 4755
End Sub

Private Sub mnuClose_Click()
Me.Hide
End Sub

Private Sub mnuCopyHelpClip_Click()
Form1.mnuCopyHelp_Click
End Sub

Public Sub shw(s As String, Optional ops As String = "")

Text1.Text = s
Me.Show 1
End Sub
