
' HEADER
'------------------------------------------------------------------'
'    Description:  Add Line Connections
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
'  - uses Repositories: GuardianConfigMgr.LineConnections.tsr
' Post-condition:
'  - Line Connections exist, as defined in config file test data
'  - Configuration Manager is running; user logged in; blank screen displayed


' START SCRIPT 

Option Explicit
Print Now & " - ACTION START: Add Line Connections"
reporter.ReportNote "ACTION started - Add Line Connections"

Dim objWindow
Set objWindow = SwfWindow("GuardianConfig_LineConnections")
Dim colParents, parentNode, strName

' get parent collection for Action
Set colParents = GetConfigNodes("/UFT/Data/TestData[@name='GuardianConfig_LineConnections']/DataSet")
If colParents.Length = 0 Then
	reporter.ReportNote "Test Data for Line Connections not found in configuration file.  Skipping process..."
	ExitAction	' abort, continue to next action
End If

If Not IsScreenDisplayed(objWindow.SwfLabel("lblTitleText"), "Line Connection") Then
	NavigateToMenu objWindow.SwfTreeView("MenuTree"), "Lines and Products;Line Connection", objWindow.SwfLabel("lblTitleText"), "Line Connection"
End If

For each parentNode in colParents	' connection
	strName = parentNode.getAttribute("name")
	print Now() & " - Set " & strName
	
	If Not IsValueInArray(Split(objWindow.SwfComboBox("cmbPickLine").GetContent, vbLF), strName, False, Null, Null) Then
		objWindow.SwfButton("btnAdd").Click	
		objWindow.SwfEdit("txtLineName").Set strName
		
		' set parent fields
		Dim colFields, fieldNode, objNode
		Set colFields = parentNode.selectNodes("Field")	' get fields
		For each fieldNode in colFields	
			SetFieldOnScreen objWindow, fieldNode.getAttribute("name"), fieldNode.getAttribute("type"), fieldNode.Text
		Next
		Set colFields = Nothing	
	End If
	
	objWindow.SwfButton("btnSave").Click
Next ' parent
Set colParents = Nothing

objWindow.SwfButton("btnClose").Click	

reporter.ReportEvent micDone, "Add Line Connections", "ACTION completed"
Print Now & " - ACTION END: Add Line Connections"

SetFieldOnScreen SwfWindow("GuardianConfig_LineConnections"), "txtMachine", "WinEdit", "hello"

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