
' HEADER
'------------------------------------------------------------------'
'    Description:  Add Notification Rules
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
'  - uses Repositories: GuardianConfigMgr.NotificationRules.tsr
' Post-condition:
'  - Notification Rules exist, as defined in config file test data
'  - Configuration Manager is running; user logged in; blank screen displayed


' START SCRIPT 

Option Explicit
Print Now & " - ACTION START: Add Notification Rules"
reporter.ReportNote "ACTION started - Add Notification Rules"

Dim objWindow
Set objWindow = SwfWindow("GuardianConfig_NotificationRules")
Dim colParents, parentNode, strName

' get parent collection for Action
Set colParents = GetConfigNodes("/UFT/Data/TestData[@name='GuardianConfig_NotificationRules']/DataSet")
If colParents.Length = 0 Then
	reporter.ReportNote "Test Data for Notification Rules not found in configuration file.  Skipping process..."
	ExitAction	' abort, continue to next action
End If

If Not IsScreenDisplayed(objWindow.SwfLabel("lblTitleText"), "Notification Rules") Then
	NavigateToMenu objWindow.SwfTreeView("MenuTree"), "Site Setup;Notification Rules", objWindow.SwfLabel("lblTitleText"), "Notification Rules"
End If

For each parentNode in colParents	' rule
	strName = parentNode.getAttribute("name")
	print Now() & " - Set " & strName & "; setDialogs=" & parentNode.getAttribute("setDialogs")

	If FindRowInDataGrid(objWindow.SwfTable("dtNotificationSettings"), "Name", strName, False)  > -1 Then	' already exists
		SelectRowInDataGrid objWindow.SwfTable("dtNotificationSettings"), "Name", strName, False			' select for update
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
		ConfigureSchema
		ConfigureFileSetting
		ConfigureCommunication
		ConfigureConversion
	End If

	objWindow.SwfButton("btnSave").Click
Next ' parent
Set colParents = Nothing

objWindow.SwfButton("btnClose").Click	

reporter.ReportEvent micDone, "Add Notification Rules", "ACTION completed"
Print Now & " - ACTION END: Add Notification Rules"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Local subroutines and functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' DESC: Configure Schema dialog
' NOTE: Opens dialog; sets values; closes dialog
Sub ConfigureSchema()
	print Now() & " -   Set Notification Schema"

	objWindow.SwfButton("btnNotificationSchema").Click	' open dialog
	Dim dlgThis
	Set dlgThis = objWindow.SwfWindow("dlgNotificationSchema")

	' set fields
	Dim colFields, fieldNode
	Set colFields = GetConfigNodes("/UFT/Data/TestData[@name='dlgNotificationSchema']/DataSet[@parent='" & strName & "']/Field")	' get fields

	For each fieldNode in colFields
		SetFieldOnScreen dlgThis, fieldNode.getAttribute("name"), fieldNode.getAttribute("type"), fieldNode.Text								
	Next
	Set colFields = Nothing

	ConfigureParameters dlgThis
	
	dlgThis.SwfButton("btnAccept").Click				' close dialog
	Set dlgThis = Nothing
End Sub

' DESC: Configure Parameters dialog
' NOTE: Opens dialog; sets values; closes dialog
Sub ConfigureParameters(ByVal dlgParent)
	print Now() & " -   Set Notification Parameters"

	dlgParent.SwfButton("btnParameters").Click	' open dialog

	Dim dlgThis
	Set dlgThis = dlgParent.SwfWindow("dlgNotificationParameters")
	
	' load all parameters
	dlgThis.SwfButton("btnAddAll").Click

	' set parameters
	Dim colParameters, parameterNode
	Set colParameters = GetConfigNodes("/UFT/Data/TestData[@name='dlgNotificationParameters']/DataSet[@parent='" & strName & "']")	' get parameters

	For each parameterNode in colParameters
		' select parameter
		SelectRowInDataGrid dlgThis.SwfTable("dtParameters"), "Name", parameterNode.getAttribute("name"), False

		' set fields
		Dim colFields, fieldNode
		Set colFields = parameterNode.selectNodes("Field")	' get fields
	
		For each fieldNode in colFields
			SetFieldOnScreen dlgThis, fieldNode.getAttribute("name"), fieldNode.getAttribute("type"), fieldNode.Text								
		Next
		Set colFields = Nothing
		
		' save parameter
		dlgThis.SwfButton("btnSave").Click
	Next ' parameter
	Set colParameters = Nothing

	dlgThis.SwfButton("btnClose").Click			' close dialog
	Set dlgThis = Nothing
End Sub

' DESC: Configure File Settings dialog
' NOTE: Opens dialog; sets values; closes dialog
Sub ConfigureFileSetting()
	print Now() & " -   Set File Settings"

	' get fields
	Dim colFields, fieldNode
	Set colFields = GetConfigNodes("/UFT/Data/TestData[@name='dlgFileConfig']/DataSet[@parent='" & strName & "']/Field")
	If colFields.Length = 0 Then
		Exit Sub	' skip if not defined
	End If

	objWindow.SwfButton("btnFileSetting").Click	' open dialog
	
	Dim dlgThis
	Set dlgThis = objWindow.SwfWindow("dlgFileConfig")

	While dlgThis.SwfTable("dtgrdFiles").RowCount > 0
		dlgThis.SwfTable("dtgrdFiles").ClickCell 0,""	' click "X" (first cell of row)
		
		If dlgThis.Dialog("dlgConfirmDelete").Exist(3) Then
			dlgThis.Dialog("dlgConfirmDelete").WinButton("btnYes").Click
		End If	
	Wend

	' set fields
	For each fieldNode in colFields
		SetFieldOnScreen dlgThis, fieldNode.getAttribute("name"), fieldNode.getAttribute("type"), fieldNode.Text								
	Next
	Set colFields = Nothing

	dlgThis.SwfButton("btnSave").Click
	dlgThis.SwfButton("btnClose").Click			' close dialog
	Set dlgThis = Nothing
End Sub

' DESC: Configure Communication dialog
' NOTE: Opens dialog; sets values; closes dialog
Sub ConfigureCommunication()
	print Now() & " -   Set Communications"
	Dim objOther, strDialogName
	Set objOther = SwfWindow("GuardianConfig_NotificationRules")

	' determine version of dialog to access
	Select Case objWindow.SwfComboBox("cmbProtocalType").GetSelection
		Case "HTTP"
			strDialogName = "dlgCommunicationHTTP"
		Case "HTTPS"
			strDialogName = "dlgCommunicationHTTPS"
		Case "FTP"
			strDialogName = "dlgCommunicationFTP"
		Case "SFTP"
			strDialogName = "dlgCommunicationSFTP"
		'Case "Web Services. Tag Serial Manager"
		'	strDialogName = ""	 ' TODO
		Case Else
			' skip File System
			Exit Sub
	End Select

	' get fields
	Dim colFields, fieldNode
	Set colFields = GetConfigNodes("/UFT/Data/TestData[@name='" & strDialogName & "']/DataSet[@parent='" & strName & "']/Field")
	If colFields.Length = 0 Then
		Set objOther = Nothing
		Exit Sub	' skip if not defined
	End If

	objWindow.SwfButton("btnCommunication").Click	' open dialog
	Dim dlgThis
	Set dlgThis = objOther.SwfWindow(strDialogName)

	' set fields
	For each fieldNode in colFields
		SetFieldOnScreen dlgThis, fieldNode.getAttribute("name"), fieldNode.getAttribute("type"), fieldNode.Text								
	Next
	Set colFields = Nothing

	dlgThis.SwfButton("btnSave").Click				' close dialog
	Set dlgThis = Nothing
End Sub

' DESC: Configure Conversion dialog
' NOTE: Opens dialog; sets values; closes dialog
Sub ConfigureConversion()
	print Now() & " -   Set File Settings skipped (NOT IMPLEMENTED)"
	' TODO:  TO BE IMPLEMENTED
End Sub

' END SCRIPT