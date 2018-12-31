
' HEADER
'------------------------------------------------------------------'
'    Description:  Add Products
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
'  - uses Repositories: GuardianConfigMgr.Products2.tsr
' Post-condition:
'  - Products exist, as defined in config file test data
'  - Configuration Manager is running; user logged in; blank screen displayed


' START SCRIPT 

Option Explicit
Print Now & " - ACTION START: Add Products"
reporter.ReportNote "ACTION started - Add Products"

Dim objWindow
Set objWindow = SwfWindow("GuardianConfig_Products")
Dim colParents, parentNode, strName, strTabName, strGrid, strKeyColumn

' get parent collection for Action
Set colParents = GetConfigNodes("/UFT/Data/TestData[@name='GuardianConfig_Products']/DataSet")
If colParents.Length = 0 Then
	reporter.ReportNote "Test Data for Products not found in configuration file.  Skipping process..."
	ExitAction	' abort, continue to next action
End If

If Not IsScreenDisplayed(objWindow.SwfLabel("lblTitleText"), "Products") Then
	NavigateToMenu objWindow.SwfTreeView("MenuTree"), "Lines and Products;Products", objWindow.SwfLabel("lblTitleText"), "Products"
End If

For each parentNode in colParents	' product
	strName = parentNode.getAttribute("name")
	strTabName = parentNode.getAttribute("tab")
	print Now() & " - Set " & strName & "; setDialogs=" & parentNode.getAttribute("setDialogs") & "; tab=" & strTabName
	
	Select Case strTabName
		Case "Manual"
			strGrid = "dtgrdProduct"
			strKeyColumn = "Product Name"
		Case "Master Data"
			strGrid = "dtgrdMD"
			strKeyColumn = "Product Name"
		Case "Master Data Templates"
			strGrid = "dtgrdTemplates"
			strKeyColumn = "Template Name"
		Case Else
			print Now() & " - ERROR - MISSING TAB NAME"
			strGrid = "missing tab name in config file"
			strKeyColumn = "missing tab name in config file"
	End Select

	objWindow.SwfTab("tbProductControl").Select strTabName

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
	
	' configure additional data (only accessible after initial save)
	If LCase(parentNode.getAttribute("setDialogs")) = "true" Then	' assume record pre-exists
		ConfigurePackagingLevels
		ConfigureMetadata
		' TODO - ConfigureLines
		' TODO - ConfigureLCV
	End If

	objWindow.SwfButton("btnSave").Click
Next ' parent
Set colParents = Nothing

objWindow.SwfButton("btnClose").Click	

reporter.ReportEvent micDone, "Add Products", "ACTION completed"
Print Now & " - ACTION END: Add Products"


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Local subroutines and functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' DESC: Configure Packaging Levels
' NOTE: Opens dialog within dialog; sets values; closes both dialogs
Sub ConfigurePackagingLevels()
	print Now() & " -   Set Packaging Levels"
	Dim colLevels, levelNode
	
	' select packaging levels
	Set colLevels = GetConfigNodes("/UFT/Data/TestData[@name='dlgPackLevels']/DataSet[@parent='" & strName & "']")	' get levels
	If colLevels.Length = 0 Then
		Exit Sub	' skip if not defined
	End If

	objWindow.SwfButton("btnPackagingLevels").Click		' open dialog
	Dim dlgLevel
	Set dlgLevel = objWindow.SwfWindow("dlgPackLevels")

	For each levelNode in colLevels
		Dim strLevel : strLevel = levelNode.getAttribute("name")
		print Now() & " -     Set " & strLevel & "; setDialogs=" & levelNode.getAttribute("setDialogs")

		If LCase(levelNode.getAttribute("clearList")) = "true" Then	' flagged to clear existing list
			print Now() & " -      Clearing list"
			Dim dlgConfig
			Set dlgConfig = dlgLevel.SwfWindow("dlgProvisioningConfig")
			While dlgLevel.SwfTable("dtgrdSPT").RowCount > 0
				dlgLevel.SwfTable("dtgrdSPT").ClickCell 0,"Type Name"
				dlgLevel.SwfButton("btnProvisioning").Click
			
				While dlgConfig.SwfTable("dtgConfigs").RowCount > 0
					dlgConfig.SwfTable("dtgConfigs").ClickCell 0,""	' click "X" (first cell of row)
					
					If dlgConfig.Dialog("dlgConfirmDelete").Exist(3) Then
						dlgConfig.Dialog("dlgConfirmDelete").WinButton("btnYes").Click
					End If	
				Wend
				dlgConfig.SwfButton("btnClose").Click			' close dialog
			
				dlgLevel.SwfTable("dtgrdSPT").ClickCell 0,""	' click "X" (first cell of row)

				If dlgLevel.Dialog("dlgConfirmDelete").Exist(3) Then
					dlgLevel.Dialog("dlgConfirmDelete").WinButton("btnYes").Click
				End If	
			Wend
		End If

		If FindRowInDataGrid(dlgLevel.SwfTable("dtgrdSPT"), "Type Name", strLevel, False)  > -1 Then	' already exists
			SelectRowInDataGrid dlgLevel.SwfTable("dtgrdSPT"), "Type Name", strLevel, False			' select for update
		Else ' add
			dlgLevel.SwfButton("btnAdd").Click
			dlgLevel.SwfComboBox("cmbMTypes").Select strLevel
		End If

		' set packlevel fields
		Dim colFields, fieldNode
		Set colFields = levelNode.selectNodes("Field")	' get fields
		For each fieldNode in colFields
			SetFieldOnScreen dlgLevel, fieldNode.getAttribute("name"), fieldNode.getAttribute("type"), fieldNode.Text								
		Next
		Set colFields = Nothing
		
		' configure additional data (only accessible after initial save)
		If LCase(levelNode.getAttribute("setDialogs")) = "true" Then	' assume record pre-exists
			ConfigureProvisioning dlgLevel, levelNode
		End If

		dlgLevel.SwfButton("btnSave").Click
	Next ' packlevel
	Set colLevels = Nothing
	
	dlgLevel.SwfButton("btnClose").Click				' close dialog
	Set dlgLevel = Nothing
End Sub

' DESC: Configure Provisioning
' NOTE: Opens dialog within dialog; sets values; closes both dialogs
Sub ConfigureProvisioning(ByVal dlgLevel, ByVal levelNode)
	print Now() & " -       Set Provisioning"

	dlgLevel.SwfButton("btnProvisioning").Click		' open dialog

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
	Set colFormats = GetConfigNodes("/UFT/Data/TestData[@name='dlgProvisioningConfig']/DataSet[@product='" & strName & "' and @packlevel='" & levelNode.getAttribute("name") & "']") ' get formats
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
	
	dlgConfig.SwfButton("btnClose").Click			' close dialog
	Set dlgConfig = Nothing
End Sub

' DESC: Configure Metadata
' NOTE: Opens dialog within dialog; sets values; closes both dialogs
Sub ConfigureMetadata()
	print Now() & " -   Set Metadata"
	Dim colData, dataNode
	
	Set colData = GetConfigNodes("/UFT/Data/TestData[@name='dlgMetadata']/DataSet[@product='" & strName & "']")	' get name/value pairs
	If colData.Length = 0 Then
		Exit Sub	' skip if not defined
	End If

	objWindow.SwfButton("btnMetadata").Click		' open dialog
	Dim dlgThis
	Set dlgThis = objWindow.SwfWindow("dlgMetadata")

	For each dataNode in colData
		Dim strKey : strKey = dataNode.getAttribute("name")
		print Now() & " -     Set " & strKey & "; clearList=" & dataNode.getAttribute("clearList")

		If LCase(dataNode.getAttribute("clearList")) = "true" Then	' flagged to clear existing list
			print Now() & " -      Clearing list"
			
			While dlgThis.SwfTable("dtgrdMetadata").RowCount > 0
				dlgThis.SwfTable("dtgrdMetadata").ClickCell 0,""	' click "X" (first cell of row)
				
				If dlgThis.Dialog("dlgMessage").Exist(3) Then
					dlgThis.Dialog("dlgMessage").WinButton("btnYes").Click
				End If	
			Wend		
		End If

		If FindRowInDataGrid(dlgThis.SwfTable("dtgrdMetadata"), "Name", strKey, False)  > -1 Then	' already exists
			SelectRowInDataGrid dlgThis.SwfTable("dtgrdMetadata"), "Name", strKey, False			' select for update
		Else ' add
			dlgThis.SwfButton("btnAdd").Click
			dlgThis.SwfEdit("txtName").Set strKey
		End If

		' set packlevel fields
		Dim colFields, fieldNode
		Set colFields = dataNode.selectNodes("Field")	' get fields
		For each fieldNode in colFields
			SetFieldOnScreen dlgThis, fieldNode.getAttribute("name"), fieldNode.getAttribute("type"), fieldNode.Text								
		Next
		Set colFields = Nothing
		
		dlgThis.SwfButton("btnSave").Click
	Next ' name/value pair
	Set colData = Nothing
	
	dlgThis.SwfButton("btnClose").Click				' close dialog
	Set dlgThis = Nothing
End Sub


' END SCRIPT