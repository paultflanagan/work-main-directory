' HEADER
'------------------------------------------------------------------'
'    Description:  Enter SPT Number - Range
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
'  - uses Repositories: GuardianConfigMgr.ManualProvisionEntry.tsr
' Post-condition:
'  - Configuration Manager is running; user logged in; blank screen displayed


' START SCRIPT

Option Explicit
Print Now & " - ACTION START: Enter_Range"
reporter.ReportNote "ACTION started - Enter SPT Number Ranges"

datatable.LocalSheet.SetCurrentRow 1
Dim strStep : strStep = datatable.Value ("Step", dtLocalsheet)
If LCase(Trim(strStep)) = "end" Then
	reporter.ReportNote "ACTION skipped - N/A"
	ExitAction	' continue to next action
End If

' Enter Range (SAP)
Dim objWindow
Set objWindow = SwfWindow("GuardianConfig_ManualProvisionEntry")

If Not IsScreenDisplayed(objWindow.SwfLabel("lblTitleText"), "Manual Provision Entry") Then
	NavigateToMenu objWindow.SwfTreeView("MenuTree"), "Manual Operations;Manual Provision Entry", objWindow.SwfLabel("lblTitleText"), "Manual Provision Entry"
End If

' process list of ranges
While LCase(Trim(strStep)) <> "end"
	Dim dtStartTime : dtStartTime = Now()
	Dim strDescription : strDescription = datatable.Value ("Description", dtLocalsheet)
	Dim strReference : strReference = datatable.Value ("Reference", dtLocalsheet)
	Dim strExpectedResult : strExpectedResult = datatable.Value ("ExpectedResult", dtLocalsheet)
	Dim strManufacturer : strManufacturer = datatable.Value ("Manufacturer", dtLocalsheet)
	Dim strProduct : strProduct = datatable.Value ("Product", dtLocalsheet)
	Dim strFormat : strFormat = datatable.Value ("Format", dtLocalsheet)
	Dim strLevel : strLevel = datatable.Value ("PackLevel", dtLocalsheet)
	Dim strStart : strStart = datatable.Value ("StartNumber", dtLocalsheet)
	Dim strEnd : strEnd = datatable.Value ("EndNumber", dtLocalsheet)
	Dim strQty : strQty = datatable.Value ("Quantity", dtLocalsheet)
	Dim strResult : strResult = datatable.Value ("Result", dtLocalsheet)

	If Not EnterSPTNumberRange(objWindow, strManufacturer, strProduct, strFormat, strLevel, strStart, strEnd, strQty, strResult) Then
		reporter.ReportEvent micFail, "Step " & strStep, "Unexpected results entering range"
		LogResult Environment("Results_File"), False, dtStartTime, Now(), "Enter_Range", strReference, strDescription, strExpectedResult
	Else
		reporter.ReportEvent micPass, "Step " & strStep, "Succeeded: " & strDescription
		LogResult Environment("Results_File"), True, dtStartTime, Now(), "Enter_Range", strReference, strDescription, strExpectedResult
	End If

	datatable.LocalSheet.SetNextRow
	strStep = datatable.Value ("Step", dtLocalsheet)	' get next value
Wend

objWindow.SwfButton("btnClose").Click

reporter.ReportEvent micDone, "Enter SPT Number Ranges", "ACTION completed"
Print Now & " - ACTION END: Enter_Range"

' END SCRIPT