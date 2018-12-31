' HEADER
'------------------------------------------------------------------'
'    Description:  Import SPT Number - Animal Health
'
'        Project:  Guardian Configuration Manager
'   Date Created:  2017 March
'         Author:  Rich Niedzwiecki
'
'  Systech International Confidential
'  © Copyright Systech International 2017
'  The source code for this program is not published or otherwise divested of its trade secrets, 
'  irrespective of what has been deposited with the U.S. Copyright Office.
'
'      Revision History
'   Date       Version   Who              Comments         
'------------------------------------------------------------------'
'
'   20170301   8.3.0     RNiedzwiecki     Initial Release
'
' 
' Pre-condition:
'  - Configuration Manager is running; user logged in; blank screen displayed
'  - Local dataSheet has steps defined
'  - requires Libraries: Common.Library.vbs, DotNet.Library.vbs, Guardian.Library.vbs
'  - uses Repositories: GuardianConfigMgr.ChinaAnimalHealth.tsr
' Post-condition:
'  - Configuration Manager is running; user logged in; blank screen displayed


' START SCRIPT

Option Explicit
Print Now & " - ACTION START: Import_AnimalHealth"
reporter.ReportNote "ACTION started - Import Animal Health Files"

datatable.LocalSheet.SetCurrentRow 1
Dim strStep : strStep = datatable.Value ("Step", dtLocalsheet)
If LCase(Trim(strStep)) = "end" Then
	reporter.ReportNote "ACTION skipped - N/A"
	ExitAction	' continue to next action
End If

' Import Animal Health
Dim objWindow, strLevel, arrLevels(), i
Set objWindow = SwfWindow("GuardianConfig_ChinaAnimalHealth")

If Not IsScreenDisplayed(objWindow.SwfLabel("lblTitleText"), "Import China Animal Health Tracing Numbers") Then
	NavigateToMenu objWindow.SwfTreeView("MenuTree"), "Manual Operations;China Animal Health Import", objWindow.SwfLabel("lblTitleText"), "Import China Animal Health Tracing Numbers"
End If

' process list of files
While LCase(Trim(strStep)) <> "end"
	Dim dtStartTime : dtStartTime = Now()
	Dim strDescription : strDescription = datatable.Value ("Description", dtLocalsheet)
	Dim strReference : strReference = datatable.Value ("Reference", dtLocalsheet)
	Dim strExpectedResult : strExpectedResult = datatable.Value ("ExpectedResult", dtLocalsheet)
	Dim strFileName : strFileName = datatable.Value ("File", dtLocalsheet)
	Dim strResult : strResult = datatable.Value ("Result", dtLocalsheet)
	
	For i = 1 To 6	' load levels
		strLevel = datatable.Value ("Level" & i, dtLocalsheet)
		If Not StringIsNullOrEmpty(strLevel) Then
			ReDim Preserve arrLevels(i-1)
			arrLevels(i-1) = strLevel
		Else
			Exit For
		End If
	Next

	If Not ImportProvisionFileAnimalHealth(objWindow, BuildPath(ProjectWorkingPath, strFileName), strResult, arrLevels) Then
		reporter.ReportEvent micFail, "Step " & strStep, "Unexpected results importing file " & strFileName
		LogResult Environment("Results_File"), False, dtStartTime, Now(), "Import_AnimalHealth", strReference, strDescription, strExpectedResult
	Else
		reporter.ReportEvent micPass, "Step " & strStep, "Succeeded: " & strDescription
		LogResult Environment("Results_File"), True, dtStartTime, Now(), "Import_AnimalHealth", strReference, strDescription, strExpectedResult
	End If

	ReDim arrLevels(-1)

	datatable.LocalSheet.SetNextRow
	strStep = datatable.Value ("Step", dtLocalsheet)	' get next value
Wend

objWindow.SwfButton("btnClose").Click

reporter.ReportEvent micDone, "Import Animal Health Files", "ACTION completed"
Print Now & " - ACTION END: Import_AnimalHealth"

' END SCRIPT