FrameWorkFolderPath = "C:\Automation\ReportFramework"

'Set WscriptObj = CreateObject("WScript.shell")
'DriverScriptFrameWorkFolderPath = WscriptObj.currentdirectory
'FrameWorkFolderPath = LEFT(DriverScriptFrameWorkFolderPath, InStrRev(DriverScriptFrameWorkFolderPath, "\") - 1)
'Set WscriptObj = Nothing

Set FSO = CreateObject("Scripting.FileSystemObject")
Set FilesObj = FSO.GetFolder(FrameWorkFolderPath &"\Test_Results").Files

If FSO.FileExists(FrameWorkFolderPath &"\Test_Results\Results.txt") Then
	Call FSO.DeleteFile(FrameWorkFolderPath &"\Test_Results\Results.txt")
End If
