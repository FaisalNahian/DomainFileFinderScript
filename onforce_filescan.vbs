Dim fso
Dim ObjOutFile
'Creating File System Object
Set fso = CreateObject("Scripting.FileSystemObject")

'Create an output file
Set ObjOutFile = fso.CreateTextFile("OutputFiles1.csv")

'Writing CSV headers
'ObjOutFile.WriteLine("Type,File Name,File Path")

'Call the GetFile function to get all files
GetFiles("C:\Dell")

'Close the output file ObjOutFile.Close
WScript.Echo("Completed")

Function GetFiles(FolderName)
On Error Resume Next
Dim ObjFolder
Dim ObjSubFolders
Dim ObjSubFolder
Dim ObjFiles
Dim ObjFile
Set ObjFolder = fso.GetFolder(FolderName)
Set ObjFiles = ObjFolder.Files

'Write all files to output files
For Each ObjFile In ObjFiles
ObjOutFile.WriteLine(ObjFile.Name & "," & ObjFile.Path)
Next

'Getting all subfolders
Set ObjSubFolders = ObjFolder.SubFolders

For Each ObjFolder In ObjSubFolders

'Writing SubFolder Name and Path
'ObjOutFile.WriteLine("Folder," & ObjFolder.Name & "," & ObjFolder.Path)

'Getting all Files from subfolder
GetFiles(ObjFolder.Path)
Next
End Function