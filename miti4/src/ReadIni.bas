Attribute VB_Name = "ReadIni"


Option Explicit
Public Const appTitle As String = "Call URL"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Function getIni(lpKeyName As String, lpDefault As String) As String
'lpBuffer is what will be read from the file.
'The 255 makes sure that only the first 255
'characters will be read. A line in your INI file
'shouldn't be that long anyway...
Dim nSize  As Integer, lpAppName As String, lpFileName As String
Dim lpBuffer  As String
lpBuffer = Space(255)
nSize = 255

'The below line reads the information from the file
'using the 6 INI file variables and stores the info
'read into the lpBuffer variable.
lpAppName = "CallUrl"
 ' "xls"
lpFileName = App.Path & "\CallUrl.ini"
getIni = GetPrivateProfileString(lpAppName, lpKeyName, lpDefault, lpBuffer, nSize, lpFileName)
getIni = Left(lpBuffer, Val(getIni))
End Function
