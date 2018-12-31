' HEADER
'------------------------------------------------------------------'
'    Description:  Manually Request Auto Provision
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
'   20170515   8.3.1     RNiedzwiecki     Correct SQL that checks for new rows in error log 
'   20170301   8.3.0     RNiedzwiecki     Initial Release
'
' 
' Pre-condition:
'  - Configuration Manager is running; user logged in; blank screen displayed
'  - Local dataSheet has steps defined
'  - requires Libraries: Common.Library.vbs, DotNet.Library.vbs, Guardian.Library.vbs
'  - uses Repositories: GuardianConfigMgr.ManualProvisionRequest.tsr
' Post-condition:
'  - Configuration Manager is running; user logged in; blank screen displayed


' START SCRIPT

Option Explicit
Print Now & " - ACTION START: Request_Numbers"
reporter.ReportNote "ACTION started - Manually Auto Provision SPT Numbers"

datatable.LocalSheet.SetCurrentRow 1
Dim strStep : strStep = datatable.Value ("Step", dtLocalsheet)
If LCase(Trim(strStep)) = "end" Then
	reporter.ReportNote "ACTION skipped - N/A"
	ExitAction	' continue to next action
End If

Dim arrOutput()

' Manual Provision Request
Dim objWindow
Set objWindow = SwfWindow("GuardianConfig_ManualProvisionRequest")

If Not IsScreenDisplayed(objWindow.SwfLabel("lblTitleText"), "Manual Provisioning Request") Then
	NavigateToMenu objWindow.SwfTreeView("MenuTree"), "Manual Operations;Manual Provisioning Request", objWindow.SwfLabel("lblTitleText"), "Manual Provisioning Request"
End If

' process list of requests
While LCase(Trim(strStep)) <> "end"
	Dim dtStartTime : dtStartTime = Now()
	Dim strDescription : strDescription = datatable.Value ("Description", dtLocalsheet)
	Dim strReference : strReference = datatable.Value ("Reference", dtLocalsheet)
	Dim strExpectedResult : strExpectedResult = datatable.Value ("ExpectedResult", dtLocalsheet)
	Dim strManufacturer : strManufacturer = datatable.Value ("Manufacturer", dtLocalsheet)
	Dim strProduct : strProduct = datatable.Value ("Product", dtLocalsheet)
	Dim strFormat : strFormat = datatable.Value ("Format", dtLocalsheet)
	Dim strLevel : strLevel = datatable.Value ("PackLevel", dtLocalsheet)
	Dim strResult : strResult = datatable.Value ("Result", dtLocalsheet)
	Dim strLastErrorId : strLastErrorId = LastErrorLogId	' note last error (if any) as baseline 

	If Not SubmitSPTNumberRequest(objWindow, strManufacturer, strProduct, strFormat, strLevel) Then
		reporter.ReportEvent micFail, "Step " & strStep, "Failed to submit request for " & strLevel
		LogResult Environment("Results_File"), False, dtStartTime, Now(), "Request_Numbers", strReference, strDescription, strExpectedResult
	Else	' submitted successfully, confirm results
		Do	
			' pause while message queue is processing (state=0-1); must wait for latest request to finish before checking on result in logs
			ExecuteSQL GetConnectionString, "SELECT COUNT(*) FROM [Guardian].[Guardian].[MessagingQueue] where SPTMessageTypeId = 0 and StateId < 2", Null, Null, arrOutput
			If CInteger(arrOutput(0,0)) = 0 Then ' nothing pending
				Exit Do
			End If
		Loop

		' check if any new error exists
		Dim strCurrentErrorId : strCurrentErrorId = LastErrorLogId
		
		If strCurrentErrorId <> strLastErrorId Then	' new error was generated
			ExecuteSQL GetConnectionString, "SELECT Description FROM [Guardian].[ErrorsLog] where ErrorId > " & strLastErrorId & " order by ErrorId", Null, Null, arrOutput	' get desc
			
			If StringIsNullOrEmpty(strResult) Then	' no error was expected
				reporter.ReportEvent micFail, "Step " & strStep, "Request of SPT Numbers failed unexpectedly with error:" & arrOutput(0,0) 		
				LogResult Environment("Results_File"), False, dtStartTime, Now(), "Request_Numbers", strReference, strDescription, strExpectedResult
			Else
				' error was expected, verify a match
				If StringStartsWith(arrOutput(0,0), strResult) Then	' result matches expected error
					reporter.ReportEvent micPass, "Step " & strStep, "Succeeded: " & strDescription
					LogResult Environment("Results_File"), True, dtStartTime, Now(), "Request_Numbers", strReference, strDescription, strExpectedResult
				Else
					reporter.ReportEvent micFail, "Step " & strStep, "Request of SPT Numbers generated unexpected result; expected '" & strResult & "' and received: " & arrOutput(0,0)
					LogResult Environment("Results_File"), False, dtStartTime, Now(), "Request_Numbers", strReference, strDescription, strExpectedResult
				End If 			
			End If ' expected error
		Else
			' no error was generated
			If StringIsNullOrEmpty(strResult) Then	' as expected
				reporter.ReportEvent micPass, "Step " & strStep, "Succeeded: " & strDescription
				LogResult Environment("Results_File"), True, dtStartTime, Now(), "Request_Numbers", strReference, strDescription, strExpectedResult
			Else	
				' but error was expected
				reporter.ReportEvent micFail, "Step " & strStep, "Request of SPT Numbers generated unexpected result; expected '" & strResult & "' and received no error"
				LogResult Environment("Results_File"), False, dtStartTime, Now(), "Request_Numbers", strReference, strDescription, strExpectedResult
			End If
		End If ' no error generated	
	End If ' successful submit

	datatable.LocalSheet.SetNextRow
	strStep = datatable.Value ("Step", dtLocalsheet)	' get next value
Wend

objWindow.SwfButton("btnClose").Click

reporter.ReportEvent micDone, "Manually Auto Provision SPT Numbers", "ACTION completed"
Print Now & " - ACTION END: Request_Numbers"

' END SCRIPT