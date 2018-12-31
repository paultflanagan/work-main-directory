Option Explicit

' PRE-REQUISITE: ERP Simulator must be running on Guardian server with custom WSEPCConfig.xml and option 'Do not change start number' enabled
' PRE-REQUISITE: UFT must be running (otherwise uncomment Launch command below)

Dim strStatus : strStatus = "Passed"
Dim strRoot : strRoot = "C:\Automation\Duplicate Check"
Dim qtApp 'As QuickTest.Application ' Declare the Application object variable
Set qtApp = CreateObject("QuickTest.Application") ' Create the Application object
qtApp.Launch ' Start UFT
qtApp.Visible = False ' Make the UFT application visible

' Set UFT run options
qtApp.Options.Run.ImageCaptureForTestResults = "OnError"
qtApp.Options.Run.RunMode = "Normal"	' [Normal or Fast]
qtApp.Options.Run.ViewResults = False

'Dim qtAutoExportResultsOpts
'Set qtAutoExportResultsOpts = qtApp.Options.Run.AutoExportReportConfig
'qtAutoExportResultsOpts.AutoExportResults = True ' export the run results at the end of each run session
'qtAutoExportResultsOpts.StepDetailsReport = True ' export the step details part of the run results at the end of each run session
'qtAutoExportResultsOpts.DataTableReport = False ' export the data table part of the run results at the end of each run session
'qtAutoExportResultsOpts.LogTrackingReport = False ' export the log tracking part of the run results at the end of each run session
'qtAutoExportResultsOpts.ScreenRecorderReport = False ' export the screen recorder part of the run results at the end of each run session
'qtAutoExportResultsOpts.SystemMonitorReport = False ' export the system monitor part of the run results at the end of each run session
'qtAutoExportResultsOpts.ExportLocation = "C:\Automation\Results" ' export the run results to the Desktop at the end of each run session
''qtAutoExportResultsOpts.UserDefinedXSL = "C:MyCustXSL.xsl" ' customized XSL file when exporting the run results data
'qtAutoExportResultsOpts.StepDetailsReportFormat = "Detailed" ' use a customized XSL file when exporting the run results data [UserDefined or Detailed or Short]
'qtAutoExportResultsOpts.ExportForFailedRunsOnly = True ' export run results only for failed runs


'WScript.echo "BEGIN " & Now() ' FOR TESTING PURPOSES ONLY

' OPTIONAL - purge all
'Dim result : result = Msgbox("Purge Data?", vbYesNo + vbQuestion, "")
'If result = vbYes Then
'	WScript.quit
'End If

''' Purge Data ''''''''''''''''''''
strStatus = RunAdvisor("PurgeData")	


''' Run GOOD lot ''''''''''''''''''''
strStatus = RunGuardian("PreLot_Good.xls")	' import numbers
'If strStatus <> "Passed" Then
'	AbortTest
'End If
strStatus = RunAdvisor("DuplicateCheckThreeLevel_GoodLot_81x") ' run lot
'If strStatus <> "Passed" Then
'	AbortTest
'End If
strStatus = RunGuardian("PostLot_Good.xls")	' disable numbers
'If strStatus <> "Passed" Then
'	AbortTest
'End If


''' Run SAFEGUARD2 lot/When Lot Starts, it should delete dupes, email should be sent   ''''''''''''''''''''
strStatus = RunGuardian("PreLot_Remove.xls")    'disable 1st Safeguard
'If strStatus <> "Passed" Then
'	AbortTest
'End If
strStatus = RunAdvisor("DuplicateCheckThreeLevel_GoodLot2ndSafeguard_81x")     'run lot, should detect dupes right away
'If strStatus <> "Passed" Then
'	AbortTest
'End If
strStatus = RunGuardian("PostLot_Remove.xls")    're-enable 1st safeguard
'If strStatus <> "Passed" Then
'	AbortTest
'End If

''' Run ACCEPT lot ''''''''''''''''''''
strStatus = RunGuardian("PreLot_Qaccept.xls")
'If strStatus <> "Passed" Then
'	AbortTest
'End If
strStatus = RunAdvisor("DuplicateCheckThreeLevel_OverrideQuarantinelLot_81x")
'If strStatus <> "Passed" and strStatus <> "Warning" Then
'	AbortTest
'End If
strStatus = RunGuardian("PostLot_Qaccept.xls")
'If strStatus <> "Passed" Then
'	AbortTest
'End If

''' Run EXCLUDE lot ''''''''''''''''''''
strStatus = RunGuardian("PreLot_ExcludeDataname.xls")
'If strStatus <> "Passed" Then
'	AbortTest
'End If
strStatus = RunAdvisor("DuplicateCheckThreeLevel_GoodLotExDataNames_81x")
'If strStatus <> "Passed" Then
'	AbortTest
'End If
strStatus = RunGuardian("PostLot_ExcludeDataname.xls")
'If strStatus <> "Passed" Then
'	AbortTest
'End If



' TODO: strStatus = RunGuardian("PreLot_PreprintedGood.xls")
' TODO: run Advisor
' TODO: strStatus = RunGuardian("PostLot_PreprintedGood.xls")

' TODO: strStatus = RunGuardian("PreprintedUsed.xls")
' TODO: preprinted ???

' TODO: strStatus = RunGuardian("PreLot_SingleItem.xls")
' TODO: run Advisor
' TODO: strStatus = RunGuardian("PostLot_SingleItem.xls")

Set qtApp = Nothing ' Release the Application object

'WScript.echo "DONE " & Now() ' FOR TESTING PURPOSES ONLY


' DESC: Aborts the Test
' NOTE: Does not log any info to result report
Sub AbortTest()
	'WScript.echo "ABORT " & Now()	' FOR TESTING PURPOSES ONLY
	WScript.quit
End Sub

' DESC: Run Test on Guardian
'  xlsFile = Path of the file containing Action parameters for the Test
' RETURN: Status of the Test result (Passed or Failed or ?)
' NOTE: Does not log any info to result report
Function PurgeGuardian()
	Dim strGuardianTest : strGuardianTest = "C:\Automation\Shared\Tests\GrdCfgMgr_PurgeData"
	Dim strLogSheet : strLogSheet = "RESULTS"
	
	PurgeGuardian = "NA"
	qtApp.Open strGuardianTest, True ' Open in read-only mode

	'Dim qtTest 'As QuickTest.Test ' Declare a Test object variable
	'Set qtTest = qtApp.Test
	'Set qtOptions = CreateObject("QuickTest.RunResultsOptions") ' Create a Results Option object
	'qtOptions.ResultsLocation = "<TempLocation>" ' Set the Results location to temporary location
	
	qtApp.Test.Run Null, True ' Run the test and wait for completion
	'PurgeGuardian = qtApp.Test.LastRunResults.Status 

End Function

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
	'If RunAdvisor = "Passed" Then	' uncomment to unload Test
	'	qtTest.Close	
	'End If
	'Set qtTest = Nothing ' Release the Test object
End Function

' DESC: Run Test on Guardian
'  xlsFile = Path of the file containing Action parameters for the Test
' RETURN: Status of the Test result (Passed or Failed or ?)
' NOTE: Does not log any info to result report
Function RunGuardian(ByVal xlsFile)
	Dim strGuardianTest : strGuardianTest = strRoot & "\" &  "FT_Provisioning_Driver"
	Dim strLogSheet : strLogSheet = "RESULTS"
	
	RunGuardian = "NA"
	qtApp.Open strGuardianTest, True ' Open in read-only mode
	Dim strLogFile : strLogFile = qtApp.Test.Environment.Value("Results_File")

	'Dim qtTest 'As QuickTest.Test ' Declare a Test object variable
	'Set qtTest = qtApp.Test
	'Set qtOptions = CreateObject("QuickTest.RunResultsOptions") ' Create a Results Option object
	'qtOptions.ResultsLocation = "<TempLocation>" ' Set the Results location to temporary location
	
	' set runtime parameters
	With qtApp.Test.DataTable
		REM ' create local results datasheet
		REM .AddSheet(strLogSheet)
		REM .GetSheet(strLogSheet).AddParameter "Timestamp", ""
		REM .GetSheet(strLogSheet).AddParameter "Passed", ""
		REM .GetSheet(strLogSheet).AddParameter "Reference", ""
		REM .GetSheet(strLogSheet).AddParameter "Description", ""
		REM .GetSheet(strLogSheet).AddParameter "Comment", ""		

		REM ' import prior results, if any
		REM If IsFileExists(strLogFile) Then
			REM .ImportSheet strLogFile, strLogSheet, strLogSheet	
		REM End If

		' import Action parameters
		.ImportSheet strGuardianTest & "\" & xlsFile, "Set_DuplicateSafeGuards", "Set_DuplicateSafeGuards"
		.ImportSheet strGuardianTest & "\" & xlsFile, "Import_Range", "Import_Range" & " [GrdCfgMgr_LoadSptNumbers]"
		.ImportSheet strGuardianTest & "\" & xlsFile, "Import_FullyRandomList", "Import_FullyRandomList" & " [GrdCfgMgr_LoadSptNumbers]"
		.ImportSheet strGuardianTest & "\" & xlsFile, "Import_AnimalHealth", "Import_AnimalHealth" & " [GrdCfgMgr_LoadSptNumbers]"
		.ImportSheet strGuardianTest & "\" & xlsFile, "Import_PartialRandomList", "Import_PartialRandomList" & " [GrdCfgMgr_LoadSptNumbers]"
		.ImportSheet strGuardianTest & "\" & xlsFile, "Import_PreprintedLabel", "Import_PreprintedLabel" & " [GrdCfgMgr_LoadSptNumbers]"
		.ImportSheet strGuardianTest & "\" & xlsFile, "Enter_Range", "Enter_Range" & " [GrdCfgMgr_LoadSptNumbers]"
		.ImportSheet strGuardianTest & "\" & xlsFile, "Request_Numbers", "Request_Numbers" & " [GrdCfgMgr_LoadSptNumbers]"
		.ImportSheet strGuardianTest & "\" & xlsFile, "Disable_Numbers", "Disable_Numbers" & " [GrdCfgMgr_LoadSptNumbers]"
	End With

	qtApp.Test.Run Null, True ' Run the test and wait for completion
	RunGuardian = qtApp.Test.LastRunResults.Status 

	REM ' export results
	REM If Len(strLogFile) > 0 Then
		REM qtApp.Test.DataTable.ExportSheet strLogFile, strLogSheet
	REM End If
	
	'If RunGuardian = "Passed" Then	' uncomment to unload Test
	'	qtTest.Close	
	'End If

	'Set qtTest = Nothing ' Release the Test object
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
