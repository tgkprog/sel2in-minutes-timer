VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "minutes timer"
   ClientHeight    =   2850
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   3975
   Icon            =   "mn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   190
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelpAbout 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   3
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3660
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Help"
      Top             =   2400
      Width           =   185
   End
   Begin VB.CommandButton cmdTogControls 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   3660
      MaskColor       =   &H00008080&
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Hide extra controls"
      Top             =   1950
      UseMaskColor    =   -1  'True
      Width           =   185
   End
   Begin VB.Timer tmrInet 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3330
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton cmdOn 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&on"
         Height          =   495
         Left            =   126
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "On Alarm after entering minutes to count down to"
         Top             =   480
         Width           =   842
      End
      Begin VB.TextBox txtMinutes 
         Height          =   375
         Left            =   285
         TabIndex        =   0
         Text            =   "25"
         ToolTipText     =   "Time in minutes, more than 60 is ok too - becomes hours"
         Top             =   11
         Width           =   525
      End
      Begin VB.TextBox txtSeconds 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Time in seconds (optional, added to minutes. more than 60 is ok too)"
         Top             =   11
         Width           =   455
      End
      Begin VB.Timer TimerFindFiles 
         Enabled         =   0   'False
         Interval        =   110
         Left            =   4320
         Top             =   720
      End
      Begin VB.Timer tmr_runOnce 
         Interval        =   200
         Left            =   4320
         Top             =   1440
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   260
         Left            =   4320
         Top             =   240
      End
      Begin VB.CommandButton cmdOff 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&off"
         Height          =   495
         Left            =   1067
         MaskColor       =   &H0080FF80&
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Off rining alarm or cancel set alarm"
         Top             =   480
         Width           =   720
      End
      Begin VB.Frame frmExtraCntrls 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0FFFF&
         Height          =   1380
         Left            =   0
         TabIndex        =   12
         Top             =   1440
         Width           =   3375
         Begin VB.CheckBox chkRepeat 
            BackColor       =   &H00C0FFFF&
            Caption         =   "repeat"
            Height          =   255
            Left            =   960
            TabIndex        =   5
            ToolTipText     =   "If checked will reset the timer after it rings/ you stop it"
            Top             =   120
            Width           =   1455
         End
         Begin VB.CheckBox chkSnd 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&sound"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   "Play Sound on alarm Click Help to know more"
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            Height          =   360
            Left            =   0
            LinkItem        =   "cmds"
            LinkTopic       =   "MimutesTimer|cmds"
            TabIndex        =   6
            ToolTipText     =   "Any text for reminder (optional)"
            Top             =   480
            Width           =   3495
         End
         Begin VB.TextBox txtShell 
            Height          =   375
            Left            =   0
            TabIndex        =   7
            Top             =   960
            Width           =   3495
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Width           =   3135
      End
      Begin VB.Label lblSearch 
         BackColor       =   &H008080FF&
         Caption         =   "Searching for sound files. Can take 3-4 minutes"
         Height          =   855
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   3240
         Top             =   600
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   2880
         ToolTipText     =   "Alarm animation"
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   0
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Alarm"
      Begin VB.Menu mnuOn 
         Caption         =   "&On"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOff 
         Caption         =   "&Off"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSncMode 
         Caption         =   "Sync mode"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSounds 
         Caption         =   "Choose S&ound"
      End
      Begin VB.Menu mnuSoundUseDef 
         Caption         =   "Use Default Sound"
      End
      Begin VB.Menu mnuSaveRem 
         Caption         =   "&Save reminder"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLoadRem 
         Caption         =   "&Load saved reminder"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSaveAllPrefsDefalut 
         Caption         =   "Save All Prefs (Default)"
      End
      Begin VB.Menu mnuLoadPrefs 
         Caption         =   "Reload all saved prefs"
      End
   End
   Begin VB.Menu mnuOther 
      Caption         =   "&Other"
      Begin VB.Menu mnuStartup 
         Caption         =   "Add to start up"
      End
      Begin VB.Menu mnuAddProgs 
         Caption         =   "Add to Start-Programs"
      End
      Begin VB.Menu mnuWebsite 
         Caption         =   "Goto &website"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuDonation 
         Caption         =   "Make donation to project"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuHideControls 
         Caption         =   "&Hide bottom controls"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuMoreCompact 
         Caption         =   "More Compact"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "Email creator"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuAlarmSoundFileName 
         Caption         =   "View Current Alarm Sound File Name"
      End
   End
   Begin VB.Menu mnuHlpMn 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpShw 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuCopyHelp 
         Caption         =   "Copy help to clipboard"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuHelpLocalHtml 
         Caption         =   "Launch Local Html Help"
      End
      Begin VB.Menu mnuHelpWebHtml 
         Caption         =   "Help on website"
      End
      Begin VB.Menu mnuUnInstall 
         Caption         =   "Uninsall"
      End
      Begin VB.Menu mnuUnsinstallAll 
         Caption         =   "Uninstall From Installer"
      End
      Begin VB.Menu mnuDebugTop 
         Caption         =   "Debug"
         Begin VB.Menu mnuDebug1 
            Caption         =   "Debug"
         End
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''D:\prog\vb\minuteAlarm\tmr3\src\..\bin\MinutesTimer_Vb6.exe
' moving from http://sourceforge.net/ to http://code.google.com/p/sel2in-minutes-timer/
Option Explicit
Public sHelpText As String
Dim dtStart As Date
Dim dtEnd As Date


Dim bUninstalling As Boolean
Dim MY_ORIG_HT As Long
Private Const APP_CAPTION As String = "Mintues Timer "
Private Const SND_FILENAME = &H20000     ' Name is a file name.
Private Const SND_ASYNC = &H1            ' Play asyncronously.
Private Declare Function PlaySound Lib "winmm.dll" _
 Alias "PlaySoundA" (ByVal lpszName As String, _
 ByVal hModule As Long, ByVal dwFlags As Long) As Long
 
Dim iStat As Long
Dim iMins As Long, iSecs As Long
Dim TempFile As String, BatchFile As String
Dim snds(3) As String
Dim iSndFileCntr  As Integer
Dim sRngFile As String


Sub alarmOnSyncCheck(bAuto As Boolean)
If bAuto Then
    dtStart = DateAdd("n", txtMinutes, dtStart)
    dtStart = DateAdd("s", txtSeconds, dtStart)
Else
    dtStart = Now
End If
Print dtEnd = DateAdd("n", txtMinutes, dtStart)
dtEnd = DateAdd("s", txtSeconds, dtEnd)
End Sub

Sub alarmOn1(bAuto As Boolean)
Timer1.Enabled = False
Timer2.Enabled = False
txtMinutes = Val(txtMinutes)
txtSeconds = Val(txtSeconds)
If (txtMinutes < 0) Then txtMinutes = 1
If (txtSeconds < 0) Then txtSeconds = 1
If txtMinutes = 0 And txtSeconds = 0 Then txtSeconds = 25

iMins = Val(txtMinutes) - 1
iSecs = Val(txtSeconds) - 1
iStat = AppStatus_ALARM_ON


Dim sMinsMsg
Dim sSecsMsg

sMinsMsg = (iMins + 1) '& ' iSecs seconds
Timer1.Interval = 60000 ' timer for mins and seconds
If iMins = -1 Then
    sMinsMsg = ""
    Timer1.Interval = 1000 'only seconds
    'iMins = 0 'bug for timer fix later with a mode int?
ElseIf iMins = 0 Then
    sMinsMsg = "1 minute "
Else
    
    sMinsMsg = (iMins + 1) & " minutes "
End If

If iSecs = -1 Then
    sSecsMsg = ""
ElseIf iSecs = 0 Then
    sSecsMsg = "1 second "
Else
    sSecsMsg = (iSecs + 1) & " seconds "
End If
Dim dt1 As Date
dt1 = Now
Label1 = "At " & Format(dt1, "hh:nn:ss") & " hrs timer for " & sMinsMsg & sSecsMsg & "started "
Dim sSep2
sSep2 = "with"
If chkSnd.Value Then
    Label1 = Label1 & "with sound "
    sSep2 = "&&"
End If
If chkRepeat.Value Then
    Label1 = Label1 & sSep2 & " repeat on"
End If
Label2 = (iMins + 1) & " " & (iSecs + 1)
Timer1.Enabled = True
Icon = Form2.Icon ' press on while ringing

'alarmOnSyncCheck False
End Sub

Sub wrFile(f As String, d As String)
Dim fs As New FileSystemObject
Dim tx As TextStream
Dim fl As File
'Set fl = fs.GetFile(f)
'Set tx = fs.OpenTextFile(f, ForWriting, True)
''set fl=fs.cr
Set tx = fs.CreateTextFile(f, True)
'Set tx = fl.OpenAsTextStream(ForWriting)
tx.WriteLine d
tx.Close
End Sub
Function crtSht(sShortcutPath As String)
Dim sExtension As String
Dim fs As New FileSystemObject
Dim oShell As New WshShell
Dim oShortcut
Set crtSht = Nothing
'sShortcutPath = InputBox("Enter path and filename of link file: ")
If sShortcutPath <> "" Then
   'sExtension = fs.GetExtensionName(sShortcutPath)
   'Select Case sExtension
      'Case "lnk"
         'Dim oShortcut As WshShortcut
         Set oShortcut = oShell.CreateShortcut(sShortcutPath)
      'Case "url"
         'Dim oURLShortcut As WshURLShortcut
        ' Set oURLShortcut = oShell.CreateShortcut(sShortcutPath)
      'Case Else
         ' user input an invalid path or filename; display an error and
         ' exit
       '  Exit Function
   'End Select
Set crtSht = oShortcut
End If
End Function

Private Sub cmdHelpAbout_Click()
mnuHelpShw_Click
    
    '& " ~ * Tip: Have XP and do not want 2 in." & vbNewLine
End Sub

Public Function getHelpText() As String
getHelpText = " ~ * Simple applicatin to get a reminder." & vbNewLine _
    & " ~ * Enter time in minutes in the first text box (whole number 1 to 32000) and click On. Info shows when the alarm was started and how many minutes." & vbNewLine _
    & " ~ * If sound is checked then plays sound files from your app folder or windows folder" & vbNewLine _
    & " ~ * Can use the second text box for any reminder text. Menu options allow you to save and load the reminder." & vbNewLine _
    & " ~ * Other menu options to visit the web site, make a donation, and email developer (via your email client)" & vbNewLine _
    & " ~ * * Picks up upto 4 sound files from application folder :" & App.Path & "\res\  to use as reminder sound. Keep copies to use same. If there are less than 4, takes the rest from the windows folder " & Environ("windir") & vbNewLine _
    & " ~ * *  Can use the last text box to run a program like  " & Environ("windir") & "\system32\notepad.exe c:\todo.txt when alarm rings." & vbNewLine _
    & " ~ *  If Repeat box is checked then the timer reset when you turn it off or it times out while 'ringing'. If you press Off when its not ringing it cancels the timer,  even if repeat is checked." & vbNewLine _
    & " ~ *  new sync mode -f on then in repeat mode time intervals are relative to clock and time started (and not when alarm is dismissed 'off')." & vbNewLine '_
End Function

Private Sub cmdOn_Click()
On Local Error GoTo errh
alarmOn1 True
alarmOnSyncCheck False
Err.Clear
If Err.Number <> 0 Then
errh:
    MsgBox "Had a problem " & Err.Description, vbExclamation, APP_CAPTION & " Err#" & Err.Number
End If
End Sub

Private Sub cmdOff_Click()
On Local Error GoTo errh
If Timer1.Enabled Then

    Label2 = "off"
End If
iStat = AppStatus_ALARM_SWITCHED_OFF
alrmDone False

Err.Clear
If Err.Number <> 0 Then
errh:
    MsgBox "Had a problem " & Err.Description, vbExclamation, APP_CAPTION & " Err#" & Err.Number
End If

End Sub

Private Sub Command3_Click()
'Me.BorderStyle = 0
'MsgBox "border sty " & Me.BorderStyle & " min " & Me.MinButton
'Form1.MinButton = False
'Form1.MaxButton = False
'form1.WhatsThisButton
End Sub



Sub fndSnd()
Dim fso As FileSystemObject
Dim fl1 As Folder, fldr2 As Folder
Dim File1 As File, file2 As File
Dim i, fnd As Integer
Set fso = New FileSystemObject
For i = 0 To 3
    If fso.FileExists(snds(i)) Then
        fnd = fnd + 1
    Else
        Exit For
    End If
Next
lblSearch.Visible = False
If fnd > UBound(snds) Then Exit Sub
'Set fl1 = fso.GetSpecialFolder(0) 'SpecialFolderConstants.WindowsFolder)
fndSnd2 Environ("windir"), fso, fnd
End Sub

Sub fndSnd2(fl As String, fso As FileSystemObject, indx As Integer)
'Dim WshShell
' Set WshShell = CreateObject("WScript.Shell")


'Set sh = crtSht(environ("temp") & "\Launch Minutues Timer.lnk")

Dim findTxt As String
findTxt = "wav"

TempFile = Environ("temp") & "\" & App.EXEName & "flagDir1.tmp"
BatchFile = Environ("temp") & "\" & App.EXEName & "find.bat"
 '/* Check If The TempFile Exists, If So, Remove It */
If Dir(TempFile, vbNormal) <> "" Then Kill TempFile

' /* Do The Same For The Batch File */
If Dir(BatchFile, vbNormal) <> "" Then Kill BatchFile
If Dir(Environ("temp") & "\dirRes343.tmp1", vbNormal) <> "" Then Kill Environ("temp") & "\dirRes343.tmp1"

' /* Open The Batch File For Writing */
Open BatchFile For Output As #1
    ' /* Write The BatchFile */
'    Print #1, "@echo off"
    'Environ("temp")
    Print #1, "" & Left(fl, 2)
    Print #1, "cd " & Chr(34) & fl & Chr(34)
    Print #1, "dir /a/s/b *." & findTxt & " >>" & Environ("temp") & "\dirRes343.tmp1"
    Print #1, "echo complete >> " & TempFile
    Print #1, "pause "
' /* All That Is Opened Must Be Closed */
Close #1
Shell BatchFile, vbHide
TimerFindFiles.Enabled = True

''Dim fl1 As Folder, fldr2 As Folder
''Dim file1 As File, file2 As File
''If indx = 4 Then Exit Sub
''For Each file1 In fl.Files
''    If Right(file1.Name, 4) = ".wav" Then
''        snds(indx) = file1.Path
''        indx = indx + 1
''    End If
''    If indx = 4 Then Exit Sub
''Next
''For Each fl1 In fl.SubFolders
''    fndSnd2 fl1, fso, indx
''    If indx = 4 Then Exit Sub
''Next
End Sub



Private Sub cmdTogControls_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Or KeyAscii = 13 Or KeyAscii = vbEnter Then
    cmdTogControls_MouseUp 1, 1, 1, 1
End If
End Sub

Private Sub cmdTogControls_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 32 Or KeyCode = 13 Or KeyCode = vbEnter Then
'    cmdTogControls_MouseUp 1, 1, 1, 1
'End If
End Sub

Private Sub cmdTogControls_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Static iMyHtBefore As Long

If Button = 1 Then
    If Not Me.Height = 2430 Then
        'compact
        Me.Height = 2430  ' Image2.Top + Image2.Height
        mnuHideControls.Caption = "Show all controls"
        'cmdTogControls.Left = 3250
        cmdTogControls.Top = 82 'cmdTogControls.Top-100' Me.Height - (2 * cmdHelpAbout.Height) - cmdTogControls.Height
    Else
        Me.Height = MY_ORIG_HT
        mnuHideControls.Caption = "Hide bottom controls, right click for more compact"
        'cmdTogControls.Left = 3600
        cmdTogControls.Top = 130
    End If
    'frmExtraCntrls.Visible = Not frmExtraCntrls.Visible
    
Else
    If Me.Height = 1365 Then
        Me.Height = iMyHtBefore
    Else
        iMyHtBefore = Me.Height
        Me.Height = 1365
        Me.Width = 3975
    End If
End If
End Sub

Private Sub Form_LinkClose()
Debug.Print "link c"
End Sub


Private Sub Form_LinkError(LinkErr As Integer)
Debug.Print "link e " & LinkErr & " "
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
Debug.Print "link Execute " & CmdStr
If (CmdStr = "a") Then
    cmdOn_Click
End If
End Sub

Private Sub Form_LinkOpen(Cancel As Integer)
Debug.Print "link open"
End Sub

Sub initHelp()
'Dim fso As FileSystemObject
Dim fl1 As Folder, fldr2 As Folder
Dim File1 As File, file2 As File
Dim i, fnd As Integer
Dim tx As TextStream
On Local Error Resume Next


If Not fso.FileExists(App.Path & "\ST6UNST.LOG") Then
    mnuUnsinstallAll.Enabled = False
End If

On Local Error GoTo errh
sHelpText = ""
''Set fso = New FileSystemObject
Set tx = fso.OpenTextFile(App.Path & "\res\help.txt", ForReading)

sHelpText = tx.ReadAll
sHelpText = Replace(sHelpText, "/--{appFolder}--/", App.Path)
sHelpText = Replace(sHelpText, "/--{windir}--/", Environ("windir"))
tx.Close
Exit Sub
errh:
If sHelpText = "" Then
    sHelpText = getHelpText
End If
End Sub

Private Sub Form_Load()
Me.mnuDebugTop.Visible = False
On Local Error Resume Next
Set fso = New FileSystemObject
Form1.LinkTopic = "MinutesTimer|cmds"
Me.Text2.LinkItem = "t2"
Me.Text2.LinkMode = 2
Form1.LinkMode = 2
Call inetInit
Call initHelp
MY_ORIG_HT = Me.Height
''D:\prog\vb\minuteAlarm\tmr3\src\..\bin\MinutesTimer_Vb6.exe
Label2.ToolTipText = "Shows time left for alarm or t/off if alarm rang or cancelled"
txtShell.ToolTipText = "Program to run on alarm like :" & Environ("windir") & "\system32\notepad.exe c:\todo.txt (optional)"
Label1 = " Loaded app At " & Format(Now, "hh:nn:ss")
Image2.ToolTipText = Image1.ToolTipText
'Form1.Icon = Form3.Icon
'Form3.Icon = Form2.Icon
'Form2.Icon = Form1.Icon
Form1.Icon = Form2.Icon
On Local Error Resume Next
Image1.Width = 800
Image1.Height = 800
Image2.Height = 800
Image2.Width = 800

Image1.Picture = LoadPicture(App.Path & ".\res\" & "tmr3.ico", , , Image1.Width, Image1.Width)
Image2.Picture = LoadPicture(App.Path & ".\res\" & "tmr3.ico")
mnuLoadPrefs_Click
End Sub

Private Sub Form_Resize()
Frame1.Left = MAX(0.45 * (Me.ScaleWidth - Frame1.Width), 5)
Frame1.Top = MAX(0.45 * (Me.ScaleHeight - Frame1.Height), 5)
End Sub

Function MAX(X As Variant, Y As Variant) As Variant
If X > Y Then
    MAX = X
Else
    MAX = Y
End If

End Function


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Function getAboutStr() As String
 getAboutStr = "By Tushar Kapila Copyright 2007-2009 tgkprog@gmail.com " & vbNewLine _
 & "Project web site http://code.google.com/p/sel2in-minutes-timer/" & vbNewLine _
 & "and  http://sel2in.com/prjs/php/p8/MinutesTimer/" & vbNewLine _
 & "Uninstall : Run the uninstall menu and simply delete all files in this apps folder" & vbNewLine _
 & "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Function


Private Sub mnuAbout_Click()
MsgBox getAboutStr, vbInformation, APP_CAPTION
End Sub

Private Sub mnuAddProgs_Click()
Dim WshShell
 Set WshShell = CreateObject("WScript.Shell")
 
On Local Error Resume Next
shortCutAdd WshShell.SpecialFolders("StartMenu") & "\Programs"
End Sub

Private Sub mnuAlarmSoundFileName_Click()
If sRngFile = "" Then
    MsgBox "You have not customized this (from alarm menu / Choose Sound) currently its a sound file found randomly on your system:""" & snds(Abs(0 Mod 4)) & """.", vbInformation, APP_CAPTION
Else
    MsgBox "Current Sound :""" & sRngFile & """.", vbInformation, APP_CAPTION
End If
End Sub

Public Sub mnuCopyHelp_Click()
On Local Error GoTo errh
Clipboard.Clear
Clipboard.SetText APP_CAPTION & vbNewLine & getHelpText & vbNewLine _
    & getAboutStr
Err.Clear
If Err.Number <> 0 Then
errh:
    MsgBox "Sorry, could not set clipboard, had a problem " & Err.Description, vbExclamation, APP_CAPTION & " Err#" & Err.Number
End If
End Sub

Private Sub mnuDebug1_Click()
MsgBox "app is production no debug values"
End Sub

Private Sub mnuDonation_Click()
On Local Error GoTo errh
Call Shell(App.Path & "\res\Minutes_Alarm_Donate_To_Project.bat", vbMinimizedFocus)
'htmlLaunch
Err.Clear
If Err.Number <> 0 Then
errh:
    MsgBox "Had a problem " & Err.Description, vbExclamation, APP_CAPTION & " Err#" & Err.Number
End If
End Sub

Private Sub mnuEmail_Click()
On Local Error GoTo errh
Call Shell(App.Path & "\res\e-Mail_Developer.bat")
'htmlLaunch App.Path & "\res\Minute_Timer.html", "mailto:TGKprog@gmail.com?subject=minutes_alarm_app.html"
Err.Clear
If Err.Number <> 0 Then
errh:
    MsgBox "Had a problem " & Err.Description, vbExclamation, APP_CAPTION & " Err#" & Err.Number
End If

'Set W = crtSht(App.Path & "\res\Minutes_Alarm_web_site.url")
End Sub


Sub htmlLaunch(sUrl, sUrl2)
If sUrl2 = "" Then sUrl2 = sUrl
Dim Dummy As String
Dim BrowserExec As String * 255
Dim RetVal As Long
Dim FileNumber As Integer
RetVal = FindExecutable(sUrl, Dummy, BrowserExec)
BrowserExec = Trim(BrowserExec)
' If an application is found, launch it!
If RetVal <= 32 Or IsEmpty(BrowserExec) Then ' Error
    MsgBox "Could not find associated Browser for """ & sUrl & """", vbExclamation, _
      "Browser Not Found"
Else
    RetVal = ShellExecute(Me.hwnd, "open", BrowserExec, _
      sUrl, Dummy, SW_SHOW)
    If RetVal <= 32 Then        ' Error
        MsgBox "Sorry could not open  " & sUrl, vbExclamation, APP_CAPTION
    End If
End If

      
End Sub
Private Sub mnuHelpLocalHtml_Click()
On Local Error Resume Next
Dim sFileNm As String
sFileNm = App.Path & "\res\Minute_Timer.html"
If fso.FileExists(sFileNm) Then
    'Shell sfileNm
    htmlLaunch sFileNm, ""
Else
    MsgBox "Cant find help file " & sFileNm & vbNewLine & "Try going to the website using other help menu", vbExclamation, APP_CAPTION
End If
End Sub

Private Sub mnuHelpShw_Click()
'cmdHelpAbout_Click
    Call Form2.shw(sHelpText & vbNewLine _
    & getAboutStr) ', vbInformation, APP_CAPTION
End Sub

Private Sub mnuHideControls_Click()
cmdTogControls_MouseUp 1, 0, 1, 1
End Sub

Private Sub mnuLoadPrefs_Click()
On Local Error Resume Next
sRngFile = GetSetting(App.EXEName, "Set", "rngFile")
chkSnd.Value = GetSetting(App.EXEName, "Set", "soundEnabled")
chkRepeat.Value = GetSetting(App.EXEName, "Set", "repeat")
mnuSncMode.Checked = GetSetting(App.EXEName, "Set", "sync", mnuSncMode.Checked)
mnuLoadRem_Click
End Sub

Private Sub mnuLoadRem_Click()
On Local Error Resume Next
Me.Text2 = GetSetting(App.EXEName, "Set", "rem", Me.Text2)
txtShell = GetSetting(App.EXEName, "Set", "shl", txtShell)
txtMinutes.Text = GetSetting(App.EXEName, "Set", "time", txtMinutes.Text)
txtSeconds.Text = GetSetting(App.EXEName, "Set", "time-seconds", txtSeconds.Text)
End Sub

Private Sub mnuMoreCompact_Click()
cmdTogControls_MouseUp 2, 1, 1, 1
End Sub

Private Sub mnuOff_Click()
On Local Error Resume Next
cmdOff_Click
End Sub

Private Sub mnuOn_Click()
On Local Error Resume Next
cmdOn_Click
End Sub

Private Sub mnuSaveAllPrefsDefalut_Click()
On Local Error GoTo errh
mnuSaveRem_Click
Call SaveSetting(App.EXEName, "Set", "rngFile", sRngFile)
Call SaveSetting(App.EXEName, "Set", "soundEnabled", chkSnd.Value)
Call SaveSetting(App.EXEName, "Set", "repeat", chkRepeat.Value)
Call SaveSetting(App.EXEName, "Set", "sync", mnuSncMode.Checked)


Err.Clear
If Err.Number <> 0 Then
errh:
    Debug.Print Err.Number & " " & Err.Description
    Resume Next
End If
End Sub

Private Sub mnuSaveRem_Click()
Call SaveSetting(App.EXEName, "Set", "time-seconds", txtSeconds.Text)
Call SaveSetting(App.EXEName, "Set", "time", txtMinutes)
Call SaveSetting(App.EXEName, "Set", "rem", Me.Text2)
Call SaveSetting(App.EXEName, "Set", "shl", Me.txtShell)
End Sub

Private Sub mnuSncMode_Click()
mnuSncMode.Checked = Not mnuSncMode.Checked
End Sub

Private Sub mnuSounds_Click()
'MsgBox "Not yet implemented", vbExclamation, APP_CAPTION
Dim s As String
s = "Wave Audio Files " & vbNullChar & "*.wav" & vbNullChar & _
              "All files" & vbNullChar & "*.*"
sRngFile = ShowOpen(s, "", "Ring file (has to be a Wave file .wav ")
End Sub

Private Sub mnuSoundUseDef_Click()
sRngFile = ""
End Sub

Private Sub mnuStartup_Click()
Dim WshShell
 Set WshShell = CreateObject("WScript.Shell")
 
On Local Error Resume Next
shortCutAdd WshShell.SpecialFolders("Startup")
'Set sh = crtSht() & "\Launch Minutues Timer.lnk")

End Sub

Sub shortCutAdd(toFld As String)
On Local Error GoTo errh

Dim WshShell
 Set WshShell = CreateObject("WScript.Shell")

Dim W As WshURLShortcut
Dim sh As WshShortcut
Set sh = crtSht(toFld & "\Launch Minutues Timer.lnk")
sh.TargetPath = App.Path & "\" & App.EXEName & ".exe"
sh.Description = "Launch the Minutes Timer Application"
sh.IconLocation = App.Path & "\" & App.EXEName & ".exe, 0"
sh.Save
Err.Clear
If Err.Number <> 0 Then
errh:
    MsgBox "Had a problem " & Err.Description, vbExclamation, APP_CAPTION & " Err#" & Err.Number
End If

End Sub


Sub delFiles(fl As Folder)
On Error Resume Next
Dim fil As File
For Each fil In fl.Files
        fil.Delete True
Next
End Sub

Sub delFolder(fl As Folder)
''On Error Resume Next
Dim fld As Folder
For Each fld In fl.SubFolders
        delFiles fld
        delFolder fld
        fld.Delete True
Next
End Sub
Private Sub mnuUnInstall_Click()
''On Local Error GoTo errh
Dim s, i, k, fl As Folder, fil As File
i = MsgBox("Uninstall all? Press Cancel to stop, Yes to remove files and registry settings, No to only remove registry entries" _
    , vbYesNoCancel, APP_CAPTION & " Uninstall")

If i = vbCancel Then Exit Sub
Set fso = New FileSystemObject
bUninstalling = True
If i = vbYes Then
    Set fl = fso.GetFolder(App.Path)
    On Local Error GoTo errHDel
    delFiles fl
    delFolder fl
    GoTo okDel
errHDel:
    Resume Next
okDel:
End If
On Local Error GoTo errh
DeleteSetting App.EXEName, "Set"
On Local Error Resume Next
DeleteSetting App.EXEName, ""

MsgBox "Removed registry entries of '" & App.EXEName & "' now simply delete app files from " & App.Path & " after the application closes" _
  & vbNewLine & "If you reopen application the registry settings will be put back again", vbInformation, APP_CAPTION
End
Err.Clear
If Err.Number <> 0 Then
errh:
    MsgBox "Try manual delete. Had a problem " & Err.Description, vbExclamation, APP_CAPTION & " Err#" & Err.Number
    Resume Next
End If
End Sub

Private Sub mnuHelpWebHtml_Click()
On Local Error Resume Next
Me.htmlLaunch "http://code.google.com/p/sel2in-minutes-timer/w/list", ""
End Sub

Private Sub mnuUnsinstallAll_Click()
''On Local Error GoTo errh
Dim s, i, k, fl As Folder, fil As File



If fso.FileExists(App.Path & "\ST6UNST.LOG") Then
i = MsgBox("Uninstall Windows Installation ? Press No to stop, Yes to continue" _
    , vbYesNo, APP_CAPTION & " Uninstall")
    If i = vbCancel Then Exit Sub
    Shell "C:\\WINDOWS\\st6unst.exe -n \""" & App.Path & "\ST6UNST.LOG""  "
    End
Else
    mnuUnInstall_Click
End If

End Sub

Private Sub mnuWebsite_Click()
On Local Error GoTo errh
'Call Shell(App.Path & "\Minutes_Alarm_Website.bat")
Me.htmlLaunch "http://sel2in.com/prjs/php/p8/MinutesTimer/", ""
Err.Clear
If Err.Number <> 0 Then
errh:
    MsgBox "Had a problem " & Err.Description, vbExclamation, APP_CAPTION & " Err#" & Err.Number
End If
End Sub



Private Sub Text2_LinkOpen(Cancel As Integer)
On Local Error Resume Next
Debug.Print "link o te"
End Sub

Private Sub Timer1_Timer()
On Local Error GoTo errh
If iStat = AppStatus_ALARM_ON And iMins > 0 Then
    If iMins > 0 Then
        iMins = iMins - 1
    End If
    Label2 = iMins & " " & iSecs
    Exit Sub
End If
If dtEnd <= Now Then
    iMins = 0
    iSecs = 0
End If
If iStat = AppStatus_ALARM_ON And (iMins = 0 Or iSecs > 0) Then
    If iMins <= 0 Then
        If iSecs > 0 Then
            If Timer1.Interval > 1000 Then
                Timer1.Interval = 1000
                Exit Sub
            Else
                iSecs = iSecs - 1
                Dim sLbl As String
                If iMins > 0 Then
                    sLbl = iMins
                Else
                    sLbl = "0"
                End If
                Label2 = sLbl & " " & iSecs
                Exit Sub
            End If
        End If
        'Timer1.Interval = 1000
    End If
    alarmOnSyncCheck True
    If txtShell <> "" Then
        On Error Resume Next
        Shell txtShell, vbNormalFocus
    End If
    On Local Error GoTo errh
    iStat = 60
    'Timer1.Interval = 1000
    Label2 = "a"
    
Else
    If iStat > -1 Then
        Timer2.Enabled = True
        iStat = iStat - 1
        If iStat Mod 2 = 0 Then
            Form1.Icon = Form2.Icon
            
        Else
            Form1.Icon = Form3.Icon
        End If
        If chkSnd.Value Then

            If sRngFile = "" Then
                Call PlaySound(snds(Abs(iStat Mod 4)), 0, SND_FILENAME Or SND_ASYNC)
            Else
                Call PlaySound(sRngFile, 0, SND_FILENAME Or SND_ASYNC)
            End If
        End If

        
    Else    'timed out
        iStat = AppStatus_ALARM_TIMED_OUT
        Timer1.Enabled = False
        Label2 = "t"
        alrmDone True
    End If
End If
Err.Clear
If Err.Number <> 0 Then
errh:
 Debug.Print Err.Description
 Resume Next
End If
 
End Sub

Sub alrmDone(bFromT As Boolean)
Icon = Form2.Icon


If chkRepeat.Value = 1 And (Timer2.Enabled Or bFromT) Then
    alarmOn1 False
Else
    Timer1.Enabled = False
End If

Timer2.Enabled = False
End Sub

Private Sub Timer2_Timer()
Static i As Integer
i = i + 1
'Image2.Visible = False
If i = 1 Then
    Image1.Left = Image1.Left - 15
    Image1.Top = Image1.Top + 15
ElseIf i = 2 Then
    Image1.Top = Image1.Top - 15
    Image1.Left = Image1.Left + 15
ElseIf i = 3 Then
    Image1.Left = Image1.Left + 15
    Image1.Top = Image1.Top + 15
ElseIf i = 4 Then
    Image1.Left = Image1.Left - 15
    Image1.Top = Image1.Top - 15
    i = 0
End If
    

End Sub



Private Sub TimerFindFiles_Timer()
Dim fso As FileSystemObject
Dim fl1 As Folder, fldr2 As Folder
Dim File1 As File, file2 As File
Dim i, fnd As Integer
Set fso = New FileSystemObject
For i = 0 To 3
    If fso.FileExists(snds(i)) Then
        fnd = fnd + 1
    Else
        Exit For
    End If
Next
lblSearch.Visible = False
If fnd > UBound(snds) Then Exit Sub

On Local Error Resume Next
DoEvents
If Dir(TempFile, vbNormal) = "" Then Exit Sub

If Dir(TempFile, vbNormal) <> "" Then Kill TempFile
If Dir(BatchFile, vbNormal) <> "" Then Kill BatchFile

Static iSndFileCntr As Integer, tmpvar
'Dim WshShell
'Set WshShell = CreateObject("WScript.Shell")

Open Environ("temp") & "\dirRes343.tmp1" For Input As #1
    ' /* Loop Until The End Of The Results File */
    Do Until EOF(1)
        ' /* Read In A Line From The File */
        Line Input #1, tmpvar
        ' /* Add It To Our Results Listbox */
        If snds(iSndFileCntr) = "" Then
            snds(iSndFileCntr) = tmpvar
            iSndFileCntr = iSndFileCntr + 1
        End If
        iSndFileCntr = iSndFileCntr + 1
        If iSndFileCntr > UBound(snds) Then Exit Do
        
    Loop
Close #1
TimerFindFiles.Enabled = False
Kill Environ("temp") & "\dirRes343.tmp1"
Static findFlag As Byte

If iSndFileCntr < UBound(snds) And findFlag = 0 Then
    findFlag = 1
    Me.fndSnd2 Environ("ProgramFiles"), Nothing, iSndFileCntr
Else
    lblSearch.Visible = False
End If



End Sub

Private Sub run_once_b()
On Local Error Resume Next
Dim s
s = GetSetting(App.EXEName, "Set", "register", "")
If (s = "") Then
    s = MsgBox("One time notify use?", vbYesNo, APP_CAPTION)
    If s = vbYes Then
        'Shell ("http://sel2in.com/prjs/php/p8/MinutesTimer/notify.php")
        notifyNet
        Call SaveSetting(App.EXEName, "Set", "register", "ok")
        If Not Inet.fso.FileExists(App.Path & "\res\a.wav_1.wav") Then
            Call Shell(App.Path & "\res\make-wav-copies.bat a.wav", vbHide)
        End If
        
    Else
        Call SaveSetting(App.EXEName, "Set", "register", "a")
    End If
    
End If
End Sub

Private Sub tmr_runOnce_Timer()

If LCase(Command$) = "/uninstall" Then
    mnuUnInstall_Click
    If bUninstalling Then
        Exit Sub
    End If
End If
Static irunOnce_state As Integer

tmr_runOnce.Enabled = False
On Local Error GoTo errh
Dim spath
Dim fso As FileSystemObject
Dim fl1 As Folder, fldr2 As Folder
Dim File1 As File, file2 As File
Dim i, fnd As Integer

Set fso = New FileSystemObject

spath = App.Path

If irunOnce_state = 0 Then
    run_once_b
    '-3 1
    fndSnd2 App.Path & "\res\", fso, fnd
    tmr_runOnce.Enabled = True
    tmr_runOnce.Interval = 15000
    irunOnce_state = irunOnce_state + 1
    Exit Sub
Dim W As WshURLShortcut
If Not fso.FileExists(App.Path & "\res\e-Mail_Developer.url") Then
    Set W = crtSht(App.Path & "\res\e-Mail_Developer.url")
    W.TargetPath = "MAILTO:TGKprog@gmail.com?subject=minutes_alarm_app"
    W.Save
End If
If Not fso.FileExists(App.Path & "\res\Minutes_Alarm_web_site.url") Then
    Set W = crtSht(App.Path & "\res\Minutes_Alarm_web_site.url")
    W.TargetPath = "http://code.google.com/p/sel2in-minutes-timer/"
    W.Save
End If
If Not fso.FileExists(App.Path & "\res\Minutes_Alarm_Donate_To_Project.url") Then
    Set W = crtSht(App.Path & "\res\Minutes_Alarm_Donate_To_Project.url")
    W.TargetPath = "http://sourceforge.net/donate/index.php?group_id=202717"
    W.Save
End If
If Not fso.FileExists(App.Path & "\res\e-Mail_Developer.bat") Then
    wrFile App.Path & "\res\e-Mail_Developer.bat", App.Path & "\e-Mail_Developer.url"
End If
If Not fso.FileExists(App.Path & "\res\Minutes_Alarm_Website.bat") Then
    wrFile App.Path & "\res\Minutes_Alarm_Website.bat", App.Path & "\res\Minutes_Alarm_web_site.url"
End If
If Not fso.FileExists(App.Path & "\res\Minutes_Alarm_Donate_To_Project.bat") Then
    wrFile App.Path & "\res\Minutes_Alarm_Donate_To_Project.bat", App.Path & "\res\Minutes_Alarm_Donate_To_Project.url"
End If

Err.Clear
If Err.Number <> 0 Then
errh:
    MsgBox "Had a problem " & Err.Description, vbExclamation, APP_CAPTION & " Err#" & Err.Number
End If

ElseIf irunOnce_state = 1 Then
    spath = Environ("WINDIR") & "\Media\Microsoft Office 2000\"
    fndSnd
    If (iSndFileCntr < 2) Then
        snds(2) = spath & "LASER.WAV"
        iSndFileCntr = iSndFileCntr + 1
    End If
    If (iSndFileCntr < 3) Then
        snds(3) = spath & "CHIMES.WAV"
        iSndFileCntr = iSndFileCntr + 1
    End If
    
    If (iSndFileCntr < 1) Then
        iSndFileCntr = iSndFileCntr + 1
        snds(0) = spath & "DRIVEBY.WAV"
    End If
    'iSndFileCntr = 1
    If (iSndFileCntr < 2) Then
        snds(1) = spath & "DRUMROLL.WAV"
     
End If

Debug.Print "stat1"
    
End If
End Sub

