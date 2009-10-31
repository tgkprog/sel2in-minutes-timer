VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4410
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   1200
      TabIndex        =   4
      Top             =   1680
      Width           =   3375
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   2880
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

