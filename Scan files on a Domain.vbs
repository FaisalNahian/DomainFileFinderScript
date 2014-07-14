''Script to find all machines in AD, search C drive of all machines for a pattern in the filename, then search any shared folders for the file and reports
'' whether it was found and whether the machines are reachable by ping.
'' Side Note -- Can not scan shares on Windows 2000 or older, it will simply skip those shares. It will still scan C drive on Windows 2000, not NT

Option Explicit
CONST HKEY_LOCAL_MACHINE = &H80000002
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Dim strResultsFile, strServerName, strErrorNumber, strErrorDescription
Dim strAge, strLatestFile, strClientName, strDNSDomain
Dim strBase, strFilter, strAttributes, strQuery
Dim objWMIService, objLocator, objResultsFile, objRootDSE, objCommand, objConnection
Dim objRecordSet
Dim strSearchName
Dim objNetwork : Set objNetwork = CreateObject("WScript.Network")
Dim strFileName : strFileName = "computers.txt"
Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim WshShell : Set WshShell = CreateObject("wscript.shell")
Dim i
Dim ii
Dim objFile, objCurrentFile, objTempList, objFS, objList, strCurrentFile2, objLogFile1, objLogFile2
Dim strComputer()
Dim strRet

'Here, you can put in the Regulare Expression pattern to use when searching.
Dim objRegEx : Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Global = True
objRegEx.IgnoreCase = True
objRegEx.Pattern = ".*pip_dvalles.*"

Dim IsFound
Dim strReply, png, strPing
Dim strList, objShare, strShare
Dim strShares()
Dim strDate, strTime, strHour, strMinute, strSeconds, Now, NowStart, ConnectTime
NowStart = Now
strResultsFile = "computers.txt"
strDate = CStr(Year(Date) * 10000 + Month(Date) * 100 + Day(Date))
strTime = Time
strHour = Hour (strTime)
strMinute = Minute (strTime)
strSeconds = Second (strTime)
Dim strFound, strNotFound
strFound = "C:\Found.txt"
strNotFound = "C:\NotFound.txt"

Set objLogFile1 = objFSO.OpenTextFile(strNotFound, ForAppending, True)
Set objLogFile2 = objFSO.OpenTextFile(strFound, ForAppending, True)

'Check for the presence of the Computer.txt file in the same folder as the script
If Not objFSO.FileExists(strFileName) Then

Set objResultsFile = objFsO.OpenTextFile (strResultsFile, ForWriting, True)
' Start getting a list of all servers from AD
' Determine DNS domain name from RootDSE object.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("DefaultNamingContext")
'Start the ADO connection
Set objCommand = CreateObject("ADODB.Command")
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection

'Set the ADO connection query strings
strBase = ""
strFilter = "(objectCategory=computer)"
strAttributes = "distinguishedName,objectCategory,name"

'Create the Query
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
objCommand.CommandText = strQuery
objCommand.Properties("Page Size") = 100
objCommand.Properties("Timeout") = 30
objCommand.Properties("Cache Results") = False
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst

'Find all computers in the domain
While Not objRecordset.EOF
ON ERROR RESUME NEXT
strServerName = objRecordset.Fields("name")
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strServerName & "\root\cimv2")
strErrorNumber = Err.Number
strErrorDescription = Err.Description

objResultsFile.WriteLine strServerName

'NEXT!
objRecordSet.MoveNext

Wend
objResultsFile.Close
WScript.Echo "All computer accounts in " & strDNSDOMAIN & " have been found." & vbCrLf & "Click OK to scan for file: " & objSearchFile
Set objFile = objFSO.OpenTextFile(strFileName, ForReading)
Else
Set objFile = objFSO.OpenTextFile(strFileName, ForReading)
End if

'start parsing through computer.txt
Do Until objFile.AtEndOfStream
isFound = FALSE
Redim Preserve strComputer(i)
strComputer(i) = objFile.ReadLine

'Ping Computers to make sure that they are reachable.
Set png = WshShell.exec("ping -n 1 " & strComputer(i))
Do Until png.Status = 1 : WScript.Sleep 10 : Loop
strPing = png.StdOut.ReadAll

'NOTE: The string being looked for in the Instr is case sensitive.
'Do not change the case of any character which appears on the
'same line as a Case InStr. AS this will result in a failure.
Select Case True
Case InStr(strPing, "Request timed out") > 1
strReply = "Request timed out"
Case InStr(strPing, "could not find host") > 1
strReply = "Host not reachable"
Case InStr(strPing, "Reply from") > 1
strReply = "Ping Successful"
End Select

' Connects to the operating system's file system
ON ERROR RESUME NEXT
Set objFS = GetObject("WinNT://" & strComputer(i) & "/LanmanServer,FileService")
objList = ""

' Loops through each share and checks for file
For Each objShare In objFS
strShare = LCase(objShare.name)
Set objTempList = WshShell.Exec("cmd /c dir /a/s/b \\" & strComputer(i) & "\" & strShare)
Do Until objTempList.StdOut.AtEndOfStream
objCurrentFile = objTempList.StdOut.ReadLine
If objRegEx.Test(objFSO.GetBaseName(objCurrentFile)) Then
strCurrentFile2 = objfSO.GetBaseName(objCurrentFile)
objLogFile2.WriteLine Now & " - The file " & objSearchFile & " was found on " & strComputer(i) & " at " & objCurrentFile
isFound = True
Else
If isFound = False Then
isFound = False
Else
isFound = True
End If
End If
Loop
objList = LCase(objShare.name) & vbCrLf & objList
Next

'Check for file on remote PC
Set objTempList = WshShell.Exec("cmd /c dir /a/s/b \\" & strComputer(i) & "c$")
Do Until objTempList.StdOut.AtEndOfStream
objCurrentFile = objTempList.StdOut.ReadLine
If objRegEx.Test(objFSO.GetBaseName(objCurrentFile)) Then
strCurrentFile2 = objfSO.GetBaseName(objCurrentFile)
objLogFile2.WriteLine Now & " - The file " & objSearchFile & " was found on " & strComputer(i) & " at " & objCurrentFile
isFound = True
Else
If isFound = False Then
isFound = False
Else
isFound = True
End If
End If
Loop

'Write to Not Found Log, if not found.
If isFound = False Then
objLogFile1.WriteLine Now & "No File matching the pattern (" & objRegEx.Pattern & ") was found on " & strComputer(i)
End If
Loop
objLogFile1.Close
objLogFile2.Close
WScript.Echo "Done scanning, LogFiles are located at:" & vbCrLf & strFound & vbCrLf & strNotFound & vbCrLf & "Click OK to finish"