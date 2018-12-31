
' HEADER
'------------------------------------------------------------------'
'    Description:  Add Manufacturers
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
'  - uses Repositories: GuardianConfigMgr.Manufacturers.tsr
' Post-condition:
'  - Manufacturers exist, as defined in config file test data
'  - Configuration Manager is running; user logged in; blank screen displayed


' START SCRIPT 

Option Explicit
Print Now & " - ACTION START: Add Manufacturers"
reporter.ReportNote "ACTION started - Add Manufacturers"

Dim objWindow
Set objWindow = SwfWindow("GuardianConfig_Manufacturers")
Dim colParents, parentNode, strName

' get parent collection for Action
Set colParents = GetConfigNodes("/UFT/Data/TestData[@name='GuardianConfig_Manufacturers']/DataSet")
If colParents.Length = 0 Then
	reporter.ReportNote "Test Data for Manufacturers not found in configuration file.  Skipping process..."
	ExitAction	' abort, continue to next action
End If

If Not IsScreenDisplayed(objWindow.SwfLabel("lblTitleText"), "Manufacturers") Then
	NavigateToMenu objWindow.SwfTreeView("MenuTree"), "Site Setup;Manufacturers", objWindow.SwfLabel("lblTitleText"), "Manufacturers"
End If

For each parentNode in colParents	' manufacturer
	strName = parentNode.getAttribute("name")
	print Now() & " - Set " & strName & "; setDialogs=" & parentNode.getAttribute("setDialogs")

	If FindRowInDataGrid(objWindow.SwfTable("dtgrdMan"), "Name", strName, False)  > -1 Then	' already exists
		SelectRowInDataGrid objWindow.SwfTable("dtgrdMan"), "Name", strName, False			' select for update
	Else ' add
		objWindow.SwfButton("btnAdd").Click	
		objWindow.SwfEdit("txtName").Set strName
	End If

	' set parent fields
	Dim colFields, fieldNode, objNode
	Set colFields = parentNode.selectNodes("Field")	' get fields
	For each fieldNode in colFields	
		SetFieldOnScreen objWindow, fieldNode.getAttribute("name"), fieldNode.getAttribute("type"), fieldNode.Text
	Next
	Set colFields = Nothing
	
	
	' configure additional data (only accessible after initial save)
	If LCase(parentNode.getAttribute("setDialogs")) = "true" Then	' assume record pre-exists
		' assign notification rules, if any
		Set objNode = parentNode.selectSingleNode("NotificationRules")
		If Not objNode Is Nothing Then
			ConfigureNotifications objNode
		End If

		ConfigureProvisioning
		ConfigureMetadata
		
		' exclude DataNames, if any
		Set objNode = parentNode.selectSingleNode("ExcludeDatanames")
		If Not objNode Is Nothing Then
			ConfigureDatanameExcludes objNode
		End If
	End If

	objWindow.SwfButton("btnSave").Click	
Next ' parent
Set colParents = Nothing

objWindow.SwfButton("btnClose").Click	

reporter.ReportEvent micDone, "Add Manufacturers", "ACTION completed"
Print Now & " - ACTION END: Add Manufacturers"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Local subroutines and functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' DESC: Configure Notification Rules list
'  objNode = node containing dialog data
' NOTE: Opens dialog; sets values; closes dialog
Sub ConfigureNotifications(ByVal objNode)
	print Now() & " -   Set Notifications"
	objWindow.SwfButton("btnSelectSequence").Click		' open dialog
	
	Dim dlgThis
	Set dlgThis = objWindow.SwfWindow("dlgNotificationSequence")

	If LCase(objNode.getAttribute("clearList")) = "true" Then	' flagged to clear existing list
		print Now() & " -    Clearing list"
		While dlgThis.SwfList("lstAssigned").GetItemsCount() > 0
			dlgThis.SwfList("lstAssigned").Select 0
			dlgThis.SwfButton("btnRemove").Click()
		Wend					
	End If
	
	' add values
	Dim colFields, fieldNode
	Set colFields = objNode.selectNodes("Field[@type='List']/Value")	' get list	
	For each fieldNode in colFields
		dlgThis.SwfList("lstAvailable").Select fieldNode.Text
		dlgThis.SwfButton("btnAdd").Click
	Next
	Set colFields = Nothing

	dlgThis.SwfButton("btnAccept").Click				' close dialog
	Set dlgThis = Nothing
End Sub

' DESC: Configure Provisioning dialog
' NOTE: Opens dialog within dialog; sets values; closes both dialogs
Sub ConfigureProvisioning()
	print Now() & " -   Set Provisioning"
	Dim colLevels, levelNode
	
	' select packaging levels
	Set colLevels = GetConfigNodes("/UFT/Data/TestData[@name='dlgSelectPackLevel']/DataSet[@parent='" & strName & "']/Field")	' get levels
	If colLevels.Length = 0 Then
		Exit Sub	' skip if not defined
	End If

	objWindow.SwfButton("btnConfiguration").Click		' open 1st dialog
	Dim dlgLevel
	Set dlgLevel = objWindow.SwfWindow("dlgSelectPackLevel")

	For each levelNode in colLevels
		dlgLevel.SwfList("lstTypes").Select levelNode.Text
		dlgLevel.SwfButton("btnConfigure").Click		' open 2nd dialog

		Dim dlgConfig
		Set dlgConfig = dlgLevel.SwfWindow("dlgProvisioningConfig")

		If LCase(levelNode.getAttribute("clearList")) = "true" Then	' flagged to clear existing list
			While dlgConfig.SwfTable("dtgConfigs").RowCount > 0
				dlgConfig.SwfTable("dtgConfigs").ClickCell 0,""	' click "X" (first cell of row)
				
				If dlgConfig.Dialog("dlgConfirmDelete").Exist(3) Then
					dlgConfig.Dialog("dlgConfirmDelete").WinButton("btnYes").Click
				End If	
			Wend
		End If

		' add SPT formats for this packlevel
		Dim colFormats, formatNode
		Set colFormats = GetConfigNodes("/UFT/Data/TestData[@name='dlgProvisioningConfig']/DataSet[@manufacturer='" & strName & "' and @packlevel='" & levelNode.Text & "']") ' get formats
		For each formatNode in colFormats
			Dim strFormat : strFormat = formatNode.getAttribute("name")
			
			If FindRowInDataGrid(dlgConfig.SwfTable("dtgConfigs"), "Format Name", strFormat, False)  > -1 Then	' already exists
				SelectRowInDataGrid dlgConfig.SwfTable("dtgConfigs"), "Format Name", strFormat, False			' select for update
			Else ' add
				dlgConfig.SwfButton("btnAdd").Click	
				dlgConfig.SwfComboBox("cmbFormats").Select strFormat
			End If

			' set format fields
			Dim colFields, fieldNode
			Set colFields = formatNode.selectNodes("Field")	' get fields
			For each fieldNode in colFields
				SetFieldOnScreen dlgConfig, fieldNode.getAttribute("name"), fieldNode.getAttribute("type"), fieldNode.Text								
			Next
			Set colFields = Nothing
			
			dlgConfig.SwfButton("btnSave").Click
		Next ' format
		Set colFormats = Nothing
		
		dlgConfig.SwfButton("btnClose").Click			' close 2nd dialog
		Set dlgConfig = Nothing
	Next ' packlevel
	Set colLevels = Nothing
	
	dlgLevel.SwfButton("btnClose").Click				' close 1st dialog
	Set dlgLevel = Nothing
End Sub

' DESC: Configure Metadata dialog
' NOTE: Opens dialog; sets values; closes dialog
Sub ConfigureMetadata
	print Now() & " -   Set Metadata skipped (NOT IMPLEMENTED)"
	' TODO: TO BE IMPLEMENTED
End Sub

' DESC: Configure DataName Exclusion dialog
'  objNode = node containing dialog data
' NOTE: Opens dialog; sets values; closes dialog
Sub ConfigureDatanameExcludes(ByVal objNode)
	print Now() & " -   Set DataName Exclusions"
	objWindow.SwfButton("btnDuplicateSerialNumber").Click	' open dialog
	
	Dim dlgThis
	Set dlgThis = objWindow.SwfWindow("dlgExcludeDatanames")

	If LCase(objNode.getAttribute("clearList")) = "true" Then	' flagged to clear existing list
		While dlgThis.SwfList("lstExcluded").GetItemsCount() > 0
			dlgThis.SwfList("lstExcluded").Select 0
			dlgThis.SwfButton("btnRight").Click()
		Wend					
	End If
	
	' add list values
	Dim colFields, fieldNode
	Set colFields = objNode.selectNodes("Field[@type='List']/Value")	' get list
	For each fieldNode in colFields
		dlgThis.SwfList("lstIncluded").Select fieldNode.Text
		dlgThis.SwfButton("btnLeft").Click
	Next
	Set colFields = Nothing

	' add non-list values
	Set colFields = objNode.selectNodes("Field[@type != 'List']")		' get fields
	For each fieldNode in colFields
		SetFieldOnScreen dlgThis, fieldNode.getAttribute("name"), fieldNode.getAttribute("type"), fieldNode.Text
	Next
	Set colFields = Nothing

	dlgThis.SwfButton("btnSave").Click						' close dialog
	Set dlgThis = Nothing
End Sub


' END SCRIPT