Attribute VB_Name = "Inet"
Option Explicit


Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_OPEN_TYPE_PROXY = 3
Const INTERNET_FLAG_RELOAD = &H80000000

Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim hOpen As Long, hFile As Long, sBuffer As String, Ret As Long

Dim cmdS As String, args(3) As String, slp As Long, lops As Long


Public fso As FileSystemObject
Dim tx As TextStream

Sub clos()
On Local Error Resume Next
InternetCloseHandle hFile
InternetCloseHandle hOpen
'Show our file
'MsgBox sBuffer

End
End Sub

Sub inetInit()
Set fso = New FileSystemObject
End Sub

Sub notifyNet()
On Local Error GoTo errh

lops = 5
slp = 500
args(0) = "http://sel2in.com/prjs/php/p8/MinutesTimer/notify.php?df3=fa&df4=g" & App.Major & "." & App.Minor & "." & App.Revision
'args(0) ="http://a.com/sel2in/prjs/php/p8/MinutesTimer/notify.php?df3=fa&df4=g" & App.Major & "." & App.Minor & "." & App.Revision
'Create an internet connection
hOpen = InternetOpen(Left(App.EXEName, Len(App.EXEName) - 4), INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
'Open the url
DoEvents
'Sleep 60000
'Set mn = New Form1
'mn.Show

hFile = InternetOpenUrl(hOpen, args(0), vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
DoEvents
read1

Err.Clear
If Err.Number <> 0 Then
errh:
End If
End Sub
''"http://www.lifegateway.com/post/c.php|6|20000"
''"http://www.lifegateway.com/post/c.php|6|20000"

Sub read1()
On Local Error GoTo errH2
Dim sBufferI As String
sBufferI = Space(1000)
sBuffer = sBufferI

Dim rtn2
rtn2 = True

'f.Timer1.Enabled = False
If lops = 0 Then
    Form1.tmrInet.Enabled = False
    tx.Close
    End
End If

Dim txLog As TextStream
Set txLog = fso.OpenTextFile(App.Path & "\CallUrl.log", ForAppending, True)
'Create a buffer for the file we're going to download
txLog.WriteLine "in process Url :" & args(0) & vbNewLine & " current lops " & lops & " wait " & (slp / 1000) & " seconds Last Read :" & Ret & vbNewLine & cmdS & vbNewLine & Format(Now, "YYYY MMM dd, hh mm ss") & vbNewLine
txLog.Close
rtn2 = 1
Ret = 1
While Ret > 0 And rtn2 = 1
    sBuffer = sBufferI
    rtn2 = InternetReadFile(hFile, sBuffer, 1000, Ret)
    ''Debug.Print sBuffer
    If (Ret > 0) Then
        tx.Write Mid(sBuffer, 1, Ret)
    End If
    
    DoEvents
    Sleep 33
    DoEvents
    'f.txtMinutes = f.txtMinutes & vbNewLine & " number read " & Ret & " rtn2 " & rtn2
    DoEvents
    'f.Refresh
    DoEvents
    Sleep 3
    DoEvents
    DoEvents
Wend
'Ret = 1

Sleep slp
lops = lops - 1

Form1.tmrInet.Enabled = True

Err.Clear
If Err.Number <> 0 Then
errh:
    Dim tx2 As TextStream
    'MsgBox "Err " & Err.Number & " " & Err.Description & vbNewLine
    Set tx2 = fso.OpenTextFile(App.Path & "\errs.log", ForAppending, True)
'Create a buffer for the file we're going to download
    tx2.WriteLine
    tx2.WriteLine Now
    Dim ss As String
    ss = "Err " & Err.Number & " " & Err.Description & vbNewLine
    Debug.Print ss
    tx2.WriteLine ss
    tx2.Close
    
    Resume Next
End If

Err.Clear
If Err.Number <> 0 Then
errH2:
End If
End Sub
