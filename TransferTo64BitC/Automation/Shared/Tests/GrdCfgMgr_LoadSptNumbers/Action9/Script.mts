' HEADER
'------------------------------------------------------------------'
'    Description:  Disable existing SPT Numbers
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
'  - uses Repositories: GuardianConfigMgr.AllocationViews.tsr, GuardianConfigMgr.PreprintedLabelImport.tsr
' Post-condition:
'  - Configuration Manager is running; user logged in; blank screen displayed


' START SCRIPT

Option Explicit
Print Now & " - ACTION START: Disable_Numbers"
reporter.ReportNote "ACTION started - Disable SPT Numbers"

Dim boolIsDisabled : boolIsDisabled = False
Dim dtStartTime : dtStartTime = Now()

' Disable Ranges
Dim objWindow
Set objWindow = SwfWindow("GuardianConfig_AllocationViews")

If Not IsScreenDisplayed(objWindow.SwfLabel("lblTitleText"), "Number Ranges Allocation") Then
	NavigateToMenu objWindow.SwfTreeView("MenuTree"), "Data Viewers;Number Ranges Allocation", objWindow.SwfLabel("lblTitleText"), "Number Ranges Allocation"
End If

datatable.LocalSheet.SetCurrentRow 1
Dim strStep : strStep = datatable.Value ("Range_Start", dtLocalsheet)

While LCase(Trim(strStep)) <> "end"		' process list of values
	DisableAvailableRange objWindow, strStep
	boolIsDisabled = True

	datatable.LocalSheet.SetNextRow
	strStep = datatable.Value ("Range_Start", dtLocalsheet)	' get next value
Wend

objWindow.SwfButton("btnClose").Click


' Disable Fully Random Lists
NavigateToMenu objWindow.SwfTreeView("MenuTree"), "Data Viewers;Number Lists Allocation", objWindow.SwfLabel("lblTitleText"), "Number Lists Allocation"

datatable.LocalSheet.SetCurrentRow 1
strStep = datatable.Value ("List_Product", dtLocalsheet)

While LCase(Trim(strStep)) <> "end"		' process list of values
	Dim strLevel : strLevel = datatable.Value ("List_PackLevel", dtLocalsheet)
	Dim strFormat : strFormat = datatable.Value ("List_Format", dtLocalsheet)
	Dim strSize : strSize = datatable.Value ("List_Size", dtLocalsheet)

	DisableAvailableList objWindow, strStep, strLevel, strFormat, strSize, True
	boolIsDisabled = True

	datatable.LocalSheet.SetNextRow
	strStep = datatable.Value ("List_Product", dtLocalsheet)	' get next value
Wend

objWindow.SwfButton("btnClose").Click


' Disable Partial Random Lists
NavigateToMenu objWindow.SwfTreeView("MenuTree"), "Data Viewers;China SFDA Allocation", objWindow.SwfLabel("lblTitleText"), "China SFDA Numbers Allocation"

datatable.LocalSheet.SetCurrentRow 1
strStep = datatable.Value ("SFDA_Resource", dtLocalsheet)

While LCase(Trim(strStep)) <> "end"		' process list of values
	Dim strStart : strStart = datatable.Value ("SFDA_Start", dtLocalsheet)

	DisableAvailableSFDA objWindow, strStep, strStart
	boolIsDisabled = True

	datatable.LocalSheet.SetNextRow
	strStep = datatable.Value ("SFDA_Resource", dtLocalsheet)	' get next value
Wend

objWindow.SwfButton("btnClose").Click


' Delete Preprinted
Set objWindow = SwfWindow("GuardianConfig_PreprintedLabel")
NavigateToMenu objWindow.SwfTreeView("MenuTree"), "Manual Operations;Preprinted Label Import", objWindow.SwfLabel("lblTitleText"), "Preprinted Label Import"

datatable.LocalSheet.SetCurrentRow 1
strStep = datatable.Value ("Preprinted_Id", dtLocalsheet)

While LCase(Trim(strStep)) <> "end"		' process list of values
	DeleteAvailablePreprinted objWindow, strStep
	boolIsDisabled = True

	datatable.LocalSheet.SetNextRow
	strStep = datatable.Value ("Preprinted_Id", dtLocalsheet)	' get next value
Wend

objWindow.SwfButton("btnClose").Click

If boolIsDisabled Then
	LogResult Environment("Results_File"), True, dtStartTime, Now(), "Disable_Numbers", Null, "Disable SPT Numbers", "Remaining SPT Numbers disabled"
Else
	reporter.ReportNote "ACTION skipped - N/A"
End If

reporter.ReportEvent micDone, "Disable SPT Numbers", "ACTION completed"
Print Now & " - ACTION END: Disable_Numbers"

' END SCRIPT