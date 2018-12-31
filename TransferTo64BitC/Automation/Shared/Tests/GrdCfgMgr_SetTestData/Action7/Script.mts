
' HEADER
'------------------------------------------------------------------'
'    Description:  Update Advanced Settings - Format Types
'
'        Project:  Guardian Configuration Manager
'   Date Created:  2017 May
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
'   20170501   8.3.0     RNiedzwiecki     Initial Release
'
' 
' Pre-condition:
'  - Configuration Manager is running; user logged in; blank screen displayed
'  - uses Global datasheet to define configuration file with test data
'  - requires Libraries: Common.Library.vbs, DotNet.Library.vbs, Guardian.Library.vbs
'  - uses Repositories: GuardianConfigMgr.AdvancedSettings.tsr
' Post-condition:
'  - Lines exist, as defined in config file test data
'  - Configuration Manager is running; user logged in; blank screen displayed


' START SCRIPT 

Option Explicit
Print Now & " - ACTION START: Update Advanced Settings Format "
reporter.ReportNote "ACTION started - Update Advanced Settings Format "

Dim objWindow
Set objWindow = SwfWindow("GuardianConfig_AdvancedSettings")
Dim colParents, parentNode, strName, strTabName, strGrid, strKeyColumn

' get parent collection for Action
Set colParents = GetConfigNodes("/UFT/Data/TestData[@name='GuardianConfig_AdvancedSettings']/DataSet")
If colParents.Length = 0 Then
	reporter.ReportNote "Test Data for Advanced Settings Foramt not found in configuration file.  Skipping process..."
	ExitAction	' abort, continue to next action
End If

If Not IsScreenDisplayed(objWindow.SwfLabel("lblTitleText"), "Lines") Then
	NavigateToMenu objWindow.SwfTreeView("MenuTree"), "General Setup;Advanced (Optional) Settings", objWindow.SwfLabel("lblTitleText"), "Advanced (Optional) Settings"
End If

For each parentNode in colParents	' line
	strName = parentNode.getAttribute("name")
	strTabName = parentNode.getAttribute("tab")
	print Now() & " - Set " & strName & "; tab=" & strTabName

	Select Case strTabName
		Case "Logistics Types"
			strGrid = "dtgrdLogistics"
			strKeyColumn = "Logistic Value"
		Case "Number Format Types"
			strGrid = "dtgrdFormats"
			strKeyColumn = "Format Name"
		Case Else
			print Now() & " - ERROR - MISSING TAB NAME"
			strGrid = "missing tab name in config file"
			strKeyColumn = "missing tab name in config file"
	End Select

	objWindow.SwfTab("tabControl").Select strTabName

	If FindRowInDataGrid(objWindow.SwfTable(strGrid), strKeyColumn, strName, False)  > -1 Then	' already exists
		SelectRowInDataGrid objWindow.SwfTable(strGrid), strKeyColumn, strName, False			' select for update
	Else ' add
		objWindow.SwfButton("btnAdd").Click	
		objWindow.SwfEdit("txtProductName").Set strName
	End If

	' set parent fields
	Dim colFields, fieldNode, objNode
	Set colFields = parentNode.selectNodes("Field")	' get fields
	For each fieldNode in colFields	
		SetFieldOnScreen objWindow, fieldNode.getAttribute("name"), fieldNode.getAttribute("type"), fieldNode.Text
	Next
	Set colFields = Nothing	
	
	objWindow.SwfButton("btnSave").Click
Next ' parent
Set colParents = Nothing

objWindow.SwfButton("btnClose").Click	

reporter.ReportEvent micDone, "Update Advanced Settings Format ", "ACTION completed"
Print Now & " - ACTION END: Update Advanced Settings Format "


' END SCRIPT