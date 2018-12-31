Option Explicit

' PRE-REQUISITE: ERP Simulator must be running on Guardian server with custom WSEPCConfig.xml and option 'Do not change start number' enabled
' PRE-REQUISITE: UFT must be running (otherwise uncomment Launch command below)

Dim strStatus : strStatus = "Passed"
' Dim bRunning : bRunning = False
Dim strRoot : strRoot = "C:\Automation\Duplicate Check"
Dim qtApp 'As QuickTest.Application ' Declare the Application object variable
Set qtApp = CreateObject("QuickTest.Application") ' Create the Application object
qtApp.Launch ' Start UFT
qtApp.Visible = False ' Make the UFT application visible

' Set UFT run options
'qtApp.Options.Run.ImageCaptureForTestResults = "OnError"
'qtApp.Options.Run.RunMode = "Normal"	' [Normal or Fast]
'qtApp.Options.Run.ViewResults = False

'Dim qtAutoExportResultsOpts
'Set qtAutoExportResultsOpts = qtApp.Options.Run.AutoExportReportConfig
'qtAutoExportResultsOpts.AutoExportResults = True ' export the run results at the end of each run session
'qtAutoExportResultsOpts.StepDetailsReport = True ' export the step details part of the run results at the end of each run session
'qtAutoExportResultsOpts.DataTableReport = False ' export the data table part of the run results at the end of each run session
'qtAutoExportResultsOpts.LogTrackingReport = False ' export the log tracking part of the run results at the end of each run session
'qtAutoExportResultsOpts.ScreenRecorderReport = False ' export the screen recorder part of the run results at the end of each run session
'qtAutoExportResultsOpts.SystemMonitorReport = False ' export the system monitor part of the run results at the end of each run session
'qtAutoExportResultsOpts.ExportLocation = "C:\Automation\Results" ' export the run results to the Desktop at the end of each run session
'qtAutoExportResultsOpts.UserDefinedXSL = "C:MyCustXSL.xsl" ' customized XSL file when exporting the run results data
'qtAutoExportResultsOpts.StepDetailsReportFormat = "Detailed" ' use a customized XSL file when exporting the run results data [UserDefined or Detailed or Short]
'qtAutoExportResultsOpts.ExportForFailedRunsOnly = True ' export run results only for failed runs


'WScript.echo "BEGIN TEST @ " & Now() ' FOR TESTING PURPOSES ONLY




' import same file as different formats

''' TAT Test ''''''''''''''''''''
'strStatus = RunAdvisor("TAT_Test")


''' Purge Data ''''''''''''''''''''
strStatus = RunAdvisor("PurgeData")


''' Run lot ''''''''''''''''''''
strStatus = RunDualFormat("PreLot_SingleItemDualFormat.xls")
'If strStatus <> "Passed" Then
'	AbortTest
'End If
strStatus = RunAdvisor("DuplicateCheckOneLevel_DualFormat_AI21")		
'If strStatus <> "Passed" Then
'	AbortTest
'End If
strStatus = RunAdvisor("DuplicateCheckOneLevel_DualFormat_GetGTIN")		
'If strStatus <> "Passed" Then
'	AbortTest
'End If
strStatus = RunDualFormat("PostLot_SingleItemDualFormat.xls")
'If strStatus <> "Passed" Then
'	AbortTest
'End If

WScript.echo "DONE @ " & Now() ' FOR TESTING PURPOSES ONLY


Set qtApp = Nothing ' Release the Application object


' DESC: Aborts the Test
' NOTE: Does not log any info to result report
Sub AbortTest()
	'WScript.echo "ABORTED @ " & Now()	' FOR TESTING PURPOSES ONLY
	WScript.quit
End Sub


' DESC: Run Test on Advisor
'  xlsFile = Path of the file containing Action parameters for the Test
' RETURN: Status of the Test result (Passed or Failed or ?)
' NOTE: Does not log any info to result report
Function RunAdvisor(ByVal strTest)
	Dim strAdvisorTest : strAdvisorTest = strRoot & "\" &  strTest
	
	RunAdvisor = "NA"
	qtApp.Open strAdvisorTest, True ' Open in read-only mode
	
	qtApp.Test.Run Null, True ' Run the test and wait for completion
	RunAdvisor = qtApp.Test.LastRunResults.Status 
End Function

' DESC: Import Numbers into Guardian using DualFormat driver
'  xlsFile = Path of the file containing Action parameters for the Test
' RETURN: Status of the Test result (Passed or Failed or ?)
Function RunDualFormat(ByVal xlsFile)
	Dim strGuardianTest : strGuardianTest = strRoot & "\" &  "FT_DualFormat_Driver"
	
	RunDualFormat = "NA"
	qtApp.Open strGuardianTest, True ' Open in read-only mode

	' set runtime parameters
	With qtApp.Test.DataTable
		' import Action parameters
		.ImportSheet strGuardianTest & "\" & xlsFile, "Import_FullyRandomList", "Import_FullyRandomList" & " [GrdCfgMgr_LoadSptNumbers]"
		.ImportSheet strGuardianTest & "\" & xlsFile, "Import_PartialRandomList", "Import_PartialRandomList" & " [GrdCfgMgr_LoadSptNumbers]"
	End With

	qtApp.Test.Run Null, True ' Run the test and wait for completion
	RunDualFormat = qtApp.Test.LastRunResults.Status 
End Function


' DESC: Copies a file from one location to another
'  strSource = The path and filename of the original file to be copied
'  strDestination = The path (and optionally filename) where the file is to be moved. If destination is a folder is must end with '\'.
'  overwrite = TRUE if any existing file at the destination is to be overwritten; otherwise FALSE to never overwrite any destination file
' NOTE: Does not log any info to result report
Sub CopyFile(ByVal strSource, ByVal strDestination, ByVal overwrite)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject") 
	
	' copy file to destination 
	If fso.FileExists(strSource) Then 
		fso.CopyFile strSource, strDestination, overwrite
	End If
	
	Set fso = Nothing
End Sub

' DESC: Verify if folder path exists
'  strFolderPath = The path of the folder to verify
' RETURN: TRUE if folder exists; otherwise FALSE if folder does not exist
' NOTE: Does not log any info to result report
Function IsFileExists(ByVal strFileName)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject") 

	IsFileExists = fso.FileExists(strFileName)

	Set fso = Nothing
End Function


' DESC: Deletes a file 
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
