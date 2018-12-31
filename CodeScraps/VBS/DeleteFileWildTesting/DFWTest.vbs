'Set Working Directory
WD = "C:\Users\paul.flanagan\MainDirectory\DeleteFileWildTesting\Files"

MsgBox(WD & "\*.txt")
DeleteFileWild(WD & "\*.txt")

' DESC: Deletes a file, after checking for its existence
'  strFilename = The path and filename of the file to be deleted
' NOTE: Does not log any info to result report
Sub DeleteFile(ByVal strFilename)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	'check for existence of file. If strFilename does not exist, fso.DeleteFile would throw a fatal error, hence the inclusion of this conditional and the omission of an else clause
	If fso.FileExists(strFilename) Then
		' delete file 
		fso.DeleteFile(strFilename)
	End If
	
	Set fso = Nothing
End Sub


' DESC: Deletes all files matching the provided glob.
'  strFilename = The glob to be deleted
' NOTE: Does not log any info to result report
Sub DeleteFileWild(ByVal strFilename)
	Dim fso, FilesObj
	Set fso = CreateObject("Scripting.FileSystemObject")
'	Set fso = CreateObject("Scripting.FileSystemProxy")
	Set FilesObj = fso.GetFile(strFilename)
'	Set FilesObj = fso.GetFiles(strFilename)
	For Each Files in FilesObj
		fso.DeleteFile(Files.Path)
	Next
	fso.
	Set fso = Nothing
'	Set FilesObj = Nothing
End Sub