Attribute VB_Name = "FileUtls2"

' Tushar Kapila tgkprog@gmail.com KSoft  2007 Copyright

Option Explicit

'------------------
'   Declarations
'------------------

' Creates an Open dialog box that lets the user specify the drive, directory, and the name
' of a file or set of files to open
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
    "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'--------------------
'   Public Methods

' Shows the open file dialog using Winwos API
Public Function ShowOpen(Filter As String, _
                         InitialDir As String, DialogTitle As String) As String
    Dim OFName As OPENFILENAME

    'Set the structure size
    OFName.lStructSize = Len(OFName)

    'Set the owner window
    OFName.hwndOwner = 0

    'Set the filter
    OFName.lpstrFilter = Filter

    'Set the maximum number of chars
    OFName.nMaxFile = 255

    'Create a buffer
    OFName.lpstrFile = Space(254)

    'Create a buffer
    OFName.lpstrFileTitle = Space$(254)

    'Set the maximum number of chars
    OFName.nMaxFileTitle = 255

    'Set the initial directory
    OFName.lpstrInitialDir = InitialDir

    'Set the dialog title
    OFName.lpstrTitle = DialogTitle

    'no extra flags
    OFName.flags = 0
    Dim ln, lp
    'Show the 'Open File' dialog
    ln = GetOpenFileName(OFName)
    If ln Then
        lp = InStr(1, OFName.lpstrFile, vbNullChar)
        ShowOpen = Mid((OFName.lpstrFile), 1, lp - 1)
    Else
        ShowOpen = ""
    End If


    
    'Me.Hide
    
'    Me.Show
'
'    ThisDocument.Activate


End Function


