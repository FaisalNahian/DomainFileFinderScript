' Scan pc's on domain for docs on c

'Declare global variables and constants
Dim fileArray, pcArray
fileArray = Array("*.doc","*.xls","*.ppt","*.docx","*.xlsx","*.pptx")
strFile = "domainpclist.txt"
strWritePath = "C:\temp\" & strFile
dim xxxxx

'find out pc names on network and write to txt file
call get_domain_pc(strWritePath, pcArray)

'loop through text file scanning each pc hard disk for 

call find_file(fileArray, pcArray)


'***********************************************************************

'VB Script to search domain pc's for files on C drive


Function get_domain_pc(strPath, pcArraylist)

'This script will list all computers on your domain

Const ADS_SCOPE_SUBTREE = 2
Const OPEN_FILE_FOR_WRITING = 2
Const ForReading = 1

Wscript.Echo "The output will be written to c:\temp\domainpclist.txt"

strDirectory = "C:\temp\"

Set objFSO1 = CreateObject("Scripting.FileSystemObject")

If objFSO1.FileExists(strPath) Then
    Set objFolder = objFSO1.GetFile(strPath)

Else
    Set objFile = objFSO1.CreateTextFile(strDirectory & strFile)
    objFile = ""

End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set textFile = fso.OpenTextFile(strPath, OPEN_FILE_FOR_WRITING)

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

' MUST CHANGE LDAP connection string to new domain
Set objCOmmand.ActiveConnection = objConnection
objCommand.CommandText = _
    "Select Name, Location from 'LDAP://DC=tekconn,DC=com' " _
        & "Where objectClass='computer'"  
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    'textFile.WriteLine(objRecordSet.Fields("Name").Value)
	'Add objRecordSet.Fields("Name").Value to pcArrayList
	
    textFile.WriteLine(objRecordSet.Fields("Name").Value)
    objRecordSet.MoveNext
Loop

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArgs = Wscript.Arguments
Set objTextFile = objFSO.OpenTextFile(strPath, ForReading)

Do Until objTextFile.AtEndOfStream
    strReg = objTextFile.Readline
Loop

WScript.Echo "Complete Finding PC Names"

End Function


Function find_file(filename, pclist)

    Dim objWMIService, colItems, objItem, strComputer
    'strComputer = "."
'external loop for pc names
  for each strComputer In pclist

   ' read text file into array
    'internal loop of file names
    '*************************************************************
	for each file in filename
    		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    		Set colItems = objWMIService.ExecQuery("SELECT * FROM CIM_DataFile WHERE FileName='" & file & "'",,48)
	next
   next

    For Each objItem in colItems
        msgbox "Found " & objItem.Name
    Next

End Function
