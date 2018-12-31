
' HEADER
'------------------------------------------------------------------'
'    Description:  Add Lines
'
'        Project:  Guardian Configuration Manager
'   Date Created:  2017 April
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
'   20170401   8.3.0     RNiedzwiecki     Initial Release
'
' 
' Pre-condition:
'  - Configuration Manager is running; user logged in; blank screen displayed
'  - uses Global datasheet to define configuration file with test data
'  - requires Libraries: Common.Library.vbs, DotNet.Library.vbs, Guardian.Library.vbs
'  - uses Repositories: GuardianConfigMgr.Lines.tsr
' Post-condition:
'  - Lines exist, as defined in config file test data
'  - Configuration Manager is running; user logged in; blank screen displayed


' START SCRIPT 

Option Explicit
Print Now & " - ACTION START: Add Lines"
reporter.ReportNote "ACTION started - Add Lines"

Dim objWindow
Set objWindow = SwfWindow("GuardianConfig_Lines")
Dim colParents, parentNode, strName, strTabName

' get parent collection for Action
Set colParents = GetConfigNodes("/UFT/Data/TestData[@name='GuardianConfig_Lines']/DataSet")
If colParents.Length = 0 Then
	reporter.ReportNote "Test Data for Lines not found in configuration file.  Skipping process..."
	ExitAction	' abort, continue to next action
End If

If Not IsScreenDisplayed(objWindow.SwfLabel("lblTitleText"), "Lines") Then
	NavigateToMenu objWindow.SwfTreeView("MenuTree"), "Lines and Products;Lines", objWindow.SwfLabel("lblTitleText"), "Lines"
End If

For each parentNode in colParents	' line
	strName = parentNode.getAttribute("name")
	strTabName = parentNode.getAttribute("tab")
	print Now() & " - Set " & strName & "; setDialogs=" & parentNode.getAttribute("setDialogs") & "; tab=" & strTabName
	
	If FindRowInDataGrid(objWindow.SwfTable("dtgrdLines"), "Line Name", strName, False)  > -1 Then	' already exists
		SelectRowInDataGrid objWindow.SwfTable("dtgrdLines"), "Line Name", strName, False			' select for update
	Else ' add
		objWindow.SwfButton("btnAdd").Click	
		objWindow.SwfEdit("txtLineName").Set strName
	End If

	objWindow.SwfTab("tabsLines").Select strTabName

	' set parent fields
	Dim colFields, fieldNode, objNode
	Set colFields = parentNode.selectNodes("Field")	' get fields
	For each fieldNode in colFields	
		SetFieldOnScreen objWindow, fieldNode.getAttribute("name"), fieldNode.getAttribute("type"), fieldNode.Text
	Next
	Set colFields = Nothing	
	
	' configure additional data (only accessible after initial save)
	If LCase(parentNode.getAttribute("setDialogs")) = "true" Then	' assume record pre-exists
		ConfigureMetadata
	End If

	objWindow.SwfButton("btnSave").Click
Next ' parent
Set colParents = Nothing

objWindow.SwfButton("btnClose").Click	

reporter.ReportEvent micDone, "Add Lines", "ACTION completed"
Print Now & " - ACTION END: Add Lines"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Local subroutines and functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' DESC: Configure Metadata dialog
' NOTE: Opens dialog; sets values; closes dialog
Sub ConfigureMetadata
	print Now() & " -   Set Metadata skipped (NOT IMPLEMENTED)"
	' TODO: TO BE IMPLEMENTED
End Sub


' END SCRIPT