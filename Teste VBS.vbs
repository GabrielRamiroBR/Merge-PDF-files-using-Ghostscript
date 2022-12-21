'Declare and Set variables
Dim strFolder
Dim objFSO
Dim wshShell


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set wshShell = WScript.CreateObject("WScript.Shell")

strUserFolder = wshShell.ExpandEnvironmentStrings("%USERPROFILE%") 'gets current user folder path
commandLine = "gswin64c.exe -dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -sOutputFile=" & strUserFolder & "\Downloads\Merge.pdf " 'saves in downloads folder
strFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)


'Set the parameters to the GhostScript merging PDF method
sParametersFiles = parametersFiles (objFSO, strFolder)

'Execute the GhostScript Merge PDf method
wshShell.Run commandLine & sParametersFiles, 0, true

'Release reference
Set objFSO = Nothing
Set wshShell = Nothing

'function that returns a string with all the pdf files to be merged in the folder (except the merged one if exist)
Function parametersFiles(ByVal objFSO, ByVal strFolder)
    pdfFiles = ""
    for each objFile in objFSO.GetFolder(strFolder).Files
        fileName = objFSO.GetFileName(objFile)
        fileType = Split(fileName, ".")(1)

        if fileType = "pdf" And not fileName = "Merge.pdf" Then
            pdfFiles = pdfFiles + " " + fileName
        End If
    Next
    parametersFiles = pdfFiles
End Function



