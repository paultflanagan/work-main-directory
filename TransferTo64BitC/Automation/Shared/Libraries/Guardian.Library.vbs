
' HEADER
'------------------------------------------------------------------'
'    Description:  Guardian General Library                        '
'                                                                  '
'        Project:  Guardian Configuration Manager Automation       '
'   Date Created:  2014 May                                        '
'         Author:  Rich Niedzwiecki                                '
'                                                                  '
' Systech International Confidential                               '
' © Copyright Systech International 2014-2017                      '
' The source code for this program is not published or otherwise   '
' divested of its trade secrets, irrespective of what has been     '
' deposited with the U.S. Copyright Office.                        '
'                                                                  '
'      Revision History                                            '
'   Date     Version   Coder          Comments                     '
'------------------------------------------------------------------'
'
'  20170515  v2.1.2    RNiedzwiecki  Corrected ImportProvisionFileAnimalHealth when successful import is unexpected
'  20170323  v2.1      RNiedzwiecki  renamed functions ImportProvisionFileSAP, ImportProvisionFileSFDA, ImportProvisionFileAnimalHealth, ImportProvisionFileReel
'                                    added functions SubmitSPTNumberRequest, DeleteAvailablePreprinted
'                                    added functions LastEventLogId, LastErrorLogId, LastProvisionLogId, LastNotificationLogId
'                                    added functions GetEventLog, GetErrorLog, GetProvisionLog, GetNotificationLog
'  20170301  v2.0      RNiedzwiecki  CODE SPLIT from Guardian Library.qfl; now contains Guardian specific functionality
'                                    added functions IsScreenDisplayed, EnterSPTNumberRange, ImportProvFileSAP, ImportProvFileSFDA, ImportProvFileAnimalHealth, ImportProvFileReel
'                                    added functions DisableAvailableRange, DisableAvailableList, DisableAvailableSFDA
'  20170131  v1.2      RNiedzwiecki  Added subroutine SetConfigFile
'                                    expanded GetConfigNode to support Environment variables (File>Settings>Environment>UserDefined)
'                                    added support for secondary configuration with metadata
'                                    added timestamps to Print debug statements
'  20160218  v1.1.21   RNiedzwiecki  Added function CopyFolder
'  20160215  v1.1.20   RNiedzwiecki  Added subroutine SetMultiTextBox
'  20160212  v1.1.19   RNiedzwiecki  Added function FindRowInDataGrid
'  20160211  v1.1.18   RNiedzwiecki  Corrected SelectRowInDataGrid to handle missing value
'  20160208  v1.1.17   RNiedzwiecki  Added function TreeViewFindNode
'  20150807  v1.1.16   RNiedzwiecki  ExecuteSQL corrected to handle no parameters
'  20150804  v1.1.15   RNiedzwiecki  Added logging info to Output tab
'  20150119  v1.1.14   RNiedzwiecki  Added function PurgeArrayColumns
'  20150116  v1.1.13   RNiedzwiecki  Added function PurgeArrayRows
'  20150113  v1.1.12   RNiedzwiecki  Expanded CheckDropDownState to support partial dropdown lists (new parameter added)
'  20150109  v1.1.11   RNiedzwiecki  Added functions GetDataGridColumnMax, GetDataGridColumnMin, GetDataGridRows
'  20141218  v1.1.10   RNiedzwiecki  Added subroutines DeleteFile and DeleteFolder
'  20141217  v1.1.9    RNiedzwiecki  Corrected ExecuteSQL to support execution of update SQL statements
'  20141210  v1.1.8    RNiedzwiecki  Added functions GetMyDocumentsPath, BuildPath, CreateFolder, IsFolderExists, IsFileExists 
'  20141124  v1.1.7    RNiedzwiecki  Added subroutine SelectRowInDataGrid
'  20141117  v1.1.6    RNiedzwiecki  Added support to execute stored procedures with one dimensional parameter array
'  20141113  v1.1.5    RNiedzwiecki  Added functions ReadFile, RegExpCount, RegExpReplacement
'  20141113  v1.1.4    RNiedzwiecki  Added subroutine TreeviewSetCheckbox; removed test code
'  20141106  v1.1.3    RNiedzwiecki  All calls to VerifyProperty use 1 sec timeout
'  20141030  v1.1.2    RNiedzwiecki  Check dialog title value
'  20141030  v1.1.1    RNiedzwiecki  Replaced RegEx support with full support for dropdown value check
'  20141027  v1.1.0    RNiedzwiecki  Added RegEx support on value check
'  20140501  v1.0.0    RNiedzwiecki  Initial version
'

' Prerequisite:
' - Project requires reference to Common.Library.vbs

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' START OF SCRIPT
Option Explicit

''' ***************************************
''' **** General Functions/Subroutines ****
''' ***************************************

Function InvokeGuardian()
Set oShell = CreateObject ("WSCript.shell")
oShell.run "cmd /C CD C:\Program Files\Systech International\Guardian SPT Config\ & GuardianSPTConfig.exe"
Set oShell = Nothing 
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Log into Guardian Configuration Manager application
'  objWindow = Top-most window repository object
'  isSQL = TRUE if authentication is SQL; otherwise FALSE if windows authentication
'  strUserId = User account id
'  strUserPwd = User account password 
'  strServerName = Name of the Guardian server
' NOTE: Does not log any info to result report
Sub DoLogin( ByVal objWindow, ByVal isSQL, ByVal strUserId, ByVal strUserPwd, ByVal strServerName)
	If isSql Then
		print Now & " Setting SQL authentication for login: " & strUserId
		objWindow.SwfComboBox("cmbAuthentication").Select "SQL Server Authentication"
		objWindow.SwfEdit("txtUsername").Set strUserId
		objWindow.SwfEdit("txtPassword").Set strUserPwd
	Else
		print Now & " Setting windows authentication for login: " & strUserId
		objWindow.SwfComboBox("cmbAuthentication").Select "Windows Authentication"
		objWindow.SwfEdit("txtPassword").Set strUserPwd
	End If
	
	objWindow.WinEdit("txtServer").Set strServerName
	objWindow.SwfButton("btnLogin").Click
	print Now & " Login completed on " & strServerName
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Select a menu (screen) and verify navigation success
'  objMenuTree = Menu tree repository object
'  strSelectedMenuPath = Menu path of screen to select  (e.g. TopMenu;SubMenu)
'  objLabel = Label repository object of selected screen
'  strTitle = Title text of selected screen
' NOTE: Does not log any info to result report
Sub NavigateToMenu(ByVal objMenuTree, ByVal strSelectedMenuPath, ByVal objLabel, ByVal strTitle)
	print Now & " - Navigate to: " & strSelectedMenuPath
	objMenuTree.Select strSelectedMenuPath			' navigate to menu
	Wait 1
	IsScreenDisplayed objLabel, strTitle			' check title to verify success
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Check if the "Confirm Exit" dialog is displayed and answer prompt as specified
'  objDialog = Confirm Exit dialog repository object
'  keepChanges = TRUE to keep changes; otherwise FALSE to ignore changes
' NOTE: Does not log any info to result report
Sub CheckConfirmExitDialog(ByVal objDialog, ByVal keepChanges)
	If objDialog.Exist(1) Then
		Dim strMsg : strMsg = objDialog.Static("lblMessage").GetROProperty("Text")

		If strMsg = "Are you sure you want to exit without saving?" Then
			If keepChanges Then
				objDialog.WinButton("btnNo").Click
			Else
				objDialog.WinButton("btnYes").Click				
			End If
			Exit Sub
		End If
		
		If strMsg = "Do you wish to save your changes before continuing?" Then
			If keepChanges Then
				objDialog.WinButton("btnYes").Click	
			Else
				objDialog.WinButton("btnNo").Click	
			End If
		End If
	End If	
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify list of displayed topmenus and submenus (does NOT include submenus not visible due to collapsed topmenu)
'  objMenuTree = Menu tree repository object
'  arrMenus = List of displayed topmenus and submenus (in sequentially displayed order)
' NOTE: Will log info to result report
Sub VerifyDisplayedMenuList(ByVal objMenuTree, ByVal arrMenus)
	Dim countVisible, countExpected, strItem, i

	countVisible = objMenuTree.GetItemsCount()
	countExpected = Ubound(arrMenus) - lbound(arrMenus) + 1

	CheckDisplayedMenuCount objMenuTree, countExpected
	
	If countVisible <> countExpected Then
		Exit sub
	End If

	For i = 0 To countVisible - 1
		strItem = objMenuTree.GetItem(i)
		
		If strItem = arrMenus(i) Then
			reporter.ReportEvent micPass, "Menu List", "Item " & i & " has the actual value '" & strItem & "'"
		Else
			reporter.ReportEvent micFail, "Menu List", "Item " & i & " has the actual value '" & strItem & "'; it was expected to be '" & arrMenus(i) & "'"
		End If		
	Next
End Sub


''' ******************************************************
''' **** Specific Functionality Functions/Subroutines ****
''' ******************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Manual Entry of SPT Number Range with specified values and verifies the expected result
'  objWindow = Top-most window repository object
'  strManufacturer = Name of Manufacturer
'  strProduct = Name of Product
'  strFormat = SPT Format
'  strLevel = Packaging level
'  strStart = Start SPT Number
'  strEnd = End of SPT Number (optional; End Number or Quantity is required)
'  strQty = Quantity of SPT Numbers (optional; End Number or Quantity is required)
'  strError = Expected error (NULL if no error is expected)
'  strError = Prefix of expected error or NULL if no error is expected  (e.g. error number)
' RETURNS: TRUE if actual result matches expected result; otherwise FALSE if unexpected result
' NOTE: Will log info to result report
Function EnterSPTNumberRange(ByVal objWindow, ByVal strManufacturer, ByVal strProduct, ByVal strFormat, ByVal strLevel, ByVal strStart, ByVal strEnd, ByVal strQty, ByVal strError)
	Dim strTestStep : strTestStep = "Enter SPT Number Range"
	Dim strResult 

	Print Now & " - Enter SPT Number range for Manufacturer=" & strManufacturer & " Product=" & strProduct & " Format=" & strFormat & " Level=" & strLevel & " Start=" & strStart & " End=" & strEnd & " Qty=" & strQty & " Error=" & strError
	reporter.ReportNote "Enter SPT Number range for Manufacturer=" & strManufacturer & " Product=" & strProduct & " Format=" & strFormat & " Level=" & strLevel & " Start=" & strStart & " End=" & strEnd & " Qty=" & strQty & " Error=" & strError
	EnterSPTNumberRange = False

	objWindow.SwfComboBox("cmbManufacturers").Select strManufacturer
	objWindow.SwfComboBox("cmbProducts").Select strProduct
	objWindow.SwfComboBox("cmbFormats").Select strFormat
	objWindow.SwfComboBox("cmbSPTObjects").Select strLevel
	objWindow.SwfEdit("txtStartSPT").Set strStart
	If Not StringIsNullOrEmpty(strEnd) Then
		objWindow.SwfEdit("txtEndSPT").Set strEnd	
		objWindow.SwfEdit("txtEndSPT").Type micReturn
	End If
	If Not StringIsNullOrEmpty(strQty) Then
		objWindow.SwfRadioButton("rdoQuantity").Set
		objWindow.SwfEdit("txtQuantity").Set strQty
		objWindow.SwfEdit("txtQuantity").Type  micReturn
	End If

	If objWindow.SwfButton("btnSave").CheckProperty("Enabled", True, 2) Then	' valid range entered
		objWindow.SwfButton("btnSave").Click
		
		If objWindow.Dialog("dlgMessage").Exist Then	' popup shown
			strResult = objWindow.Dialog("dlgMessage").Static("lblMessage").GetROProperty("Text")	' get result
			
			If StringIsNullOrEmpty(strError) Then	' no error expected
				If InStr(strResult, "successfull") > 0 Then
					reporter.ReportEvent micPass, strTestStep, "Manual Entry of SPT Number Range succeeded"
					EnterSPTNumberRange = True									
				Else
					reporter.ReportEvent micFail, strTestStep, "Manual Entry of SPT Number Range failed; received error: " & strResult					
				End If
			Else 	' error expected
				If StringStartsWith(strResult, strError) Then	' result matches expected error
					reporter.ReportEvent micPass, strTestStep, "Manual Entry of SPT Number Range generated expected error '" & strError & "'"								
					EnterSPTNumberRange = True
				Else
					reporter.ReportEvent micFail, strTestStep, "Manual Entry of SPT Number Range generated unexpected result; expected '" & strError & "' and received: " & strResult
				End If 
			End If

			' close result message
			If InStr(strResult, "successfull") > 0 Then
				objWindow.Dialog("dlgMessage").WinButton("btnOK").Click
			Else
				objWindow.Dialog("dlgMessage").WinButton("btnNo").Click
			End If		
		End If
	Else
		reporter.ReportEvent  micFail, strTestStep, "Invalid SPT Number range entered"
	End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Imports a specified SAP provisioning file and verifies the expected result
'  objWindow = Top-most window repository object
'  strFile = Path to SAP provisioning import file
'  strError = Prefix of expected error or NULL if no error is expected  (e.g. error number)
' RETURNS: TRUE if actual result matches expected result; otherwise FALSE if unexpected result
' NOTE: Will log info to result report
Function ImportProvisionFileSAP(ByVal objWindow, ByVal strFile, ByVal strError)
	Dim strTestStep : strTestStep = "Import SAP file"
	Dim strResult 
	
	Print Now & " - Importing SAP file " & strFile
	reporter.ReportNote "Importing SAP file " & strFile
	ImportProvisionFileSAP = False

	objWindow.SwfButton("btnImport").Click
	
	If objWindow.Dialog("dlgBrowseFile").Exist Then	' file open dialog shown
		' select file
		objWindow.Dialog("dlgBrowseFile").WinEdit("txtName").Set strFile
		Wait(1)
		objWindow.Dialog("dlgBrowseFile").WinButton("btnOpen").Click
		
		If objWindow.SwfWindow("dlgPreview").Exist Then	' preview dialog shown
			' start import
			objWindow.SwfWindow("dlgPreview").SwfButton("btnOK").Click
			
			Do	' wait for import to end
				If objWindow.Dialog("dlgMessage").Exist(10) Then
					Print Now & " - waiting for prompt"
					Exit Do
				End If
			Loop	' indefinitely
			
			' verify correct result
			If objWindow.Dialog("dlgMessage").Exist Then	' popup shown
				strResult = objWindow.Dialog("dlgMessage").Static("lblMessage").GetROProperty("Text") ' get result
				
				If StringIsNullOrEmpty(strError) Then	' no error expected
					If InStr(strResult, "successfull") > 0 Then
						reporter.ReportEvent micPass, strTestStep, "Import of SAP file succeeded"
						ImportProvisionFileSAP = True									
					Else
						reporter.ReportEvent micFail, strTestStep, "Import of SAP file failed; received error: " & strResult					
					End If
				Else 	' error expected
					If StringStartsWith(strResult, strError) Then	' result matches expected error
						reporter.ReportEvent micPass, strTestStep, "Import of SAP file generated expected error '" & strError & "'"								
						ImportProvisionFileSAP = True
					Else
						reporter.ReportEvent micFail, strTestStep, "Import of SAP file generated unexpected result; expected '" & strError & "' and received: " & strResult
					End If 
				End If
				
				' close result message
				If InStr(strResult, "successfull") > 0 Then
					objWindow.Dialog("dlgMessage").WinButton("btnOK").Click
				Else
					objWindow.Dialog("dlgMessage").WinButton("btnNo").Click
				End If
			Else
				reporter.ReportEvent  micFail, strTestStep, "Failed to get prompt"
			End If ' message dialog
		Else
			reporter.ReportEvent  micFail, strTestStep, "Failed to open preview dialog"		
		End If ' preview dialog
	Else	
		reporter.ReportEvent  micFail, strTestStep, "Failed to open browser dialog"
	End If ' file open dialog
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Imports a specified China SFDA provisioning file and verifies the expected result
'  objWindow = Top-most window repository object
'  strFile = Path to SFDA provisioning import file
'  strError = Prefix of expected error or NULL if no error is expected  (e.g. error number)
' RETURNS: TRUE if actual result matches expected result; otherwise FALSE if unexpected result
' NOTE: Will log info to result report
Function ImportProvisionFileSFDA(ByVal objWindow, ByVal strFile, ByVal strError)
	Dim strTestStep : strTestStep = "Import China SFDA file"
	Dim strResult
	
	Print Now & " - Importing SFDA file " & strFile
	reporter.ReportNote "Importing SFDA file " & strFile
	ImportProvisionFileSFDA = False

	objWindow.SwfButton("btnBrowse").Click
	If objWindow.Dialog("dlgBrowseFile").Exist Then	' file open dialog shown
		' select file
		objWindow.Dialog("dlgBrowseFile").WinEdit("txtName").Set strFile
		Wait(1)
		objWindow.Dialog("dlgBrowseFile").WinButton("btnOpen").Click
		
		If objWindow.SwfButton("btnImport").CheckProperty("Enabled", True) Then	' matching product found
			' start import
			objWindow.SwfButton("btnImport").Click
			
			Do	' wait for import to end
				If objWindow.Dialog("dlgMessage").Exist(10) Then
					Print Now & " - waiting for prompt"
					Exit Do
				End If
			Loop	' indefinitely
			
			' verify correct result
			If objWindow.Dialog("dlgMessage").Exist Then	' popup shown
				strResult = objWindow.Dialog("dlgMessage").Static("lblMessage").GetROProperty("Text")	' get result

				If StringIsNullOrEmpty(strError) Then	' no error expected
					If InStr(strResult, "successfull") > 0 Then
						reporter.ReportEvent micPass, strTestStep, "Import of SFDA file succeeded"
						ImportProvisionFileSFDA = True									
					Else
						reporter.ReportEvent micFail, strTestStep, "Import of SFDA file failed; received error: " & strResult					
					End If
				Else 	' error expected
					If StringStartsWith(strResult, strError) Then	' result matches expected error
						reporter.ReportEvent micPass, strTestStep, "Import of SFDA file generated expected error '" & strError & "'"								
						ImportProvisionFileSFDA = True
					Else
						reporter.ReportEvent micFail, strTestStep, "Import of SFDA file generated unexpected result; expected '" & strError & "' and received: " & strResult
					End If 
				End If

				' close result message
				If InStr(strResult, "successfull") > 0 Then
					objWindow.Dialog("dlgMessage").WinButton("btnOK").Click
				Else
					objWindow.Dialog("dlgMessage").WinButton("btnNo").Click
				End If
			End IF ' message dialog
		Else
			reporter.ReportEvent  micFail, strTestStep, "No product match found"
		End If ' product match found
	Else	
		reporter.ReportEvent  micFail, strTestStep, "Failed to open browser dialog"
	End If ' file open dialog
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Imports a specified China Animal Health provisioning file and verifies the expected result
'  strFile = Path to China Animal Health provisioning import file
'  strError = Prefix of expected error or NULL if no error is expected  (e.g. error number)
'  arrLevels = Array of packaging levels to be imported [syntax: <level> : <capacity>] (e.g. Pouch : 0)
' RETURNS: TRUE if actual result matches expected result; otherwise FALSE if unexpected result
' NOTE: Will log info to result report
Function ImportProvisionFileAnimalHealth(ByVal objWindow, ByVal strFile, ByVal strError, ByVal arrLevels)
	Dim strTestStep : strTestStep = "Import Animal Health file"
	Dim strResult, i
	
	Print Now & " - Importing Animal Health file " & strFile
	reporter.ReportNote "Importing Animal Health file " & strFile
	ImportProvisionFileAnimalHealth = True

	objWindow.SwfButton("btnBrowseZip").Click
	If objWindow.Dialog("dlgBrowseFile").Exist Then	' file open dialog shown
		' select file
		objWindow.Dialog("dlgBrowseFile").WinEdit("txtName").Set strFile
		Wait(1)
		objWindow.Dialog("dlgBrowseFile").WinButton("btnOpen").Click
		
		For i = 0 To Ubound(arrLevels)	' select each packlevel
			objWindow.SwfList("clbProductLevels").Select arrLevels(i)
		Next	
		
		If objWindow.SwfButton("btnImport").CheckProperty("Enabled", True) Then	' matching product found
			' start import
			objWindow.SwfButton("btnImport").Click
					
			If StringIsNullOrEmpty(strError) Then	' no error is expected
				For i = 0 To Ubound(arrLevels)		' each level a packlevel file is imported
					Do	' wait for import of packlevel file to end
						Print Now & " - waiting for prompt"
						If objWindow.Dialog("dlgMessage").Exist(10) Then	' got a popup message (a.k.a. completed)
							Exit Do
						End If
					Loop	' indefinitely
	
					' verify correct result
					If objWindow.Dialog("dlgMessage").Exist Then	' popup shown
						strResult = objWindow.Dialog("dlgMessage").Static("lblMessage").GetROProperty("Text")	' get result
						
						If InStr(strResult, "successfull") > 0 Then
							reporter.ReportEvent micPass, strTestStep, "Import of Animal Health file succeeded; " & strResult
						Else
							reporter.ReportEvent micFail, strTestStep, "Import of Animal Health file failed; received error: " & strResult					
							ImportProvisionFileAnimalHealth = False
						End If
					
						' close result message
						If InStr(strResult, "successfull") > 0 Then
							objWindow.Dialog("dlgMessage").WinButton("btnOK").Click
						Else
							objWindow.Dialog("dlgMessage").WinButton("btnNo").Click
						End If
						If ImportProvisionFileAnimalHealth = False Then
							Exit For
						End If
					End IF
				Next ' file
			Else 	' error is expected
				Do	' wait for import of packlevel file to end
					Print Now & " - waiting for prompt"
					If objWindow.Dialog("dlgMessage").Exist(10) Then	' got a popup message (a.k.a. completed)
						Exit Do
					End If
				Loop	' indefinitely

				' verify correct result
				If objWindow.Dialog("dlgMessage").Exist Then	' popup shown
					strResult = objWindow.Dialog("dlgMessage").Static("lblMessage").GetROProperty("Text")	' get result
					
					If StringStartsWith(strResult, strError) Then	' result matches expected error
						reporter.ReportEvent micPass, strTestStep, "Import of Animal Health generated expected error '" & strError & "'"								
					Else
						reporter.ReportEvent micFail, strTestStep, "Import of Animal Health file generated unexpected result; expected '" & strError & "' and received: " & strResult
						ImportProvisionFileAnimalHealth = False
					End If 
				
					' close result message
					If InStr(strResult, "successfull") > 0 Then
						objWindow.Dialog("dlgMessage").WinButton("btnOK").Click
					Else
						objWindow.Dialog("dlgMessage").WinButton("btnNo").Click
					End If
				End If ' message dialog
				
				If (ImportProvisionFileAnimalHealth = False) And (InStr(strResult, "successfull") > 0) Then	' failed to get expected error at the first level
					' ignore remaining packlevel imports
					For i = 1 To Ubound(arrLevels)		' each remaining level a packlevel file is imported
						Do	' wait for import of packlevel file to end
							Print Now & " - waiting for prompt"
							If objWindow.Dialog("dlgMessage").Exist(10) Then	' got a popup message (a.k.a. completed)
								Exit Do
							End If
						Loop	' indefinitely
						If objWindow.Dialog("dlgMessage").Exist Then	' popup shown
							strResult = objWindow.Dialog("dlgMessage").Static("lblMessage").GetROProperty("Text")	' get result						
							reporter.ReportEvent micFail, strTestStep, "Import of Animal Health file continued"
						
							' close result message
							If InStr(strResult, "successfull") > 0 Then
								objWindow.Dialog("dlgMessage").WinButton("btnOK").Click
							Else
								objWindow.Dialog("dlgMessage").WinButton("btnNo").Click
							End If
						End If ' message dialog
					Next ' file
				End If					
			End If ' expected error		
		Else
			reporter.ReportEvent  micFail, strTestStep, "No product match found"
			ImportProvisionFileAnimalHealth = False
		End If ' product match
	Else	
		reporter.ReportEvent  micFail, strTestStep, "Failed to open browser dialog"
		ImportProvisionFileAnimalHealth = False
	End If ' file open dialog
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Imports a specified Preprinted Label provisioning file and verifies the expected result
'  objWindow = Top-most window repository object
'  strFile = Path to Preprinted Label provisioning import file
'  strError = Prefix of expected error or NULL if no error is expected  (e.g. error number)
' RETURNS: TRUE if actual result matches expected result; otherwise FALSE if unexpected result
' NOTE: Will log info to result report
Function ImportProvisionFileReel(ByVal objWindow, ByVal strFile, ByVal strError)
	Dim strTestStep : strTestStep = "Import Preprinted Label file"
	Dim strResult 
	
	Print Now & " - Importing Preprinted Label file " & strFile
	reporter.ReportNote "Importing Preprinted Label file " & strFile
	ImportProvisionFileReel = False

	objWindow.SwfButton("btnImport").Click
	
	If objWindow.Dialog("dlgBrowseFile").Exist Then	' file open dialog shown
		' select file
		objWindow.Dialog("dlgBrowseFile").WinEdit("txtName").Set strFile
		Wait(1)
		objWindow.Dialog("dlgBrowseFile").WinObject("btnOpen").Click
	
		Do	' wait for import to end
			If objWindow.Dialog("dlgMessage").Exist(10) Then
				Print Now & " - waiting for prompt"
				Exit Do
			End If
		Loop	' indefinitely
		
		' verify correct result
		If objWindow.Dialog("dlgMessage").Exist Then	' popup shown
			strResult = objWindow.Dialog("dlgMessage").Static("lblMessage").GetROProperty("Text") ' get result
			
			If StringIsNullOrEmpty(strError) Then	' no error expected
				If InStr(strResult, "completed") > 0 Then
					reporter.ReportEvent micPass, strTestStep, "Import of Preprinted Label file succeeded"
					ImportProvisionFileReel = True									
				Else
					reporter.ReportEvent micFail, strTestStep, "Import of Preprinted Label file failed; received error: " & strResult					
				End If
			Else 	' error expected
				If StringStartsWith(strResult, strError) Then	' result matches expected error
					reporter.ReportEvent micPass, strTestStep, "Import of Preprinted Label file generated expected error '" & strError & "'"								
					ImportProvisionFileReel = True
				Else
					reporter.ReportEvent micFail, strTestStep, "Import of Preprinted Label file generated unexpected result; expected '" & strError & "' and received: " & strResult
				End If 
			End If
			
			' close result message
			If InStr(strResult, "completed") > 0 Then
				objWindow.Dialog("dlgMessage").WinButton("btnOK").Click
			Else
				objWindow.Dialog("dlgMessage").WinButton("btnNo").Click
			End If
		Else
			reporter.ReportEvent  micFail, strTestStep, "Failed to get prompt"
		End If ' message dialog
	Else	
		reporter.ReportEvent  micFail, strTestStep, "Failed to open browser dialog"
	End If ' file open dialog
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Submit request for SPT Numbers
'  objWindow = Top-most window repository object
'  strManufacturer = Name of Manufacturer
'  strProduct = Name of Product
'  strFormat = SPT Format to request
'  strLevel = Packaging level to request
' RETURNS: TRUE if successfully submitted; otherwise FALSE
' NOTE: Will log info to result report
Function SubmitSPTNumberRequest(ByVal objWindow, ByVal strManufacturer, ByVal strProduct, ByVal strFormat, ByVal strLevel)
	Dim strTestStep : strTestStep = "Submit SPT Number Request"
	Dim strResult, i

	Print Now & " - Request SPT Numbers for Manufacturer=" & strManufacturer & " Product=" & strProduct & " Format=" & strFormat & " Level=" & strLevel
	reporter.ReportNote "Request SPT Numbers for Manufacturer=" & strManufacturer & " Product=" & strProduct & " Format=" & strFormat & " Level=" & strLevel
	SubmitSPTNumberRequest = False

	objWindow.SwfComboBox("cmbManufacturers").Select strManufacturer
	objWindow.SwfComboBox("cmbProducts").Select strProduct
	objWindow.SwfComboBox("cmbFormats").Select strFormat

	For i = 0 To SwfWindow("GuardianConfig_ManualProvisionRequest").SwfList("lstSPTObjects").GetItemsCount - 1	' clear list
		SwfWindow("GuardianConfig_ManualProvisionRequest").SwfList("lstSPTObjects").Object.SetItemChecked i, False
	Next
	objWindow.SwfList("lstSPTObjects").Select strLevel
	
	objWindow.SwfButton("btnSubmit").Click

	If objWindow.Dialog("dlgMessage").Exist Then	' popup shown
		objWindow.Dialog("dlgMessage").WinButton("btnOK").Click
		reporter.ReportEvent micPass, strTestStep, "Submitted request for "	& strLevel
		SubmitSPTNumberRequest = True
	Else
		reporter.ReportEvent  micFail, strTestStep, "Failed to submit request for " & strLevel
	End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Disables a range of SPT Numbers
'  objWindow = Top-most window repository object with Availability tab displayed
'  strStart = Start SPT Number of range
' RETURNS: TRUE if found and disabled; otherwise FALSE if not found or disabled
' NOTE: Will log info to result report
Function DisableAvailableRange(ByVal objWindow, ByVal strStart)
	Print Now & " - Disable Range " & strStart
	reporter.ReportNote "Disable Range with Start=" & strStart
	
	DisableAvailableRange = False
	Dim objDataGrid 
	Set objDataGrid = objWindow.SwfTable("dtgrdRange")

	If FindRowInDataGrid(objDataGrid, "Start Number", strStart, False) > -1 Then	' match on primary key
		' select and disable
		SelectRowInDataGrid objDataGrid, "Start Number", strStart, False
	
		objWindow.SwfButton("btnDisable").Click
		If objWindow.Dialog("dlgConfirmExit").Exist(2) Then
			objWindow.Dialog("dlgConfirmExit").WinButton("btnYes").Click
			reporter.ReportEvent micPass, "Disable SPT Number Range", "Disabled range starting with " & strStart
			DisableAvailableRange = True
		End If ' dialog
	End If ' match on primary
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Disables a list of SPT Numbers
'  objWindow = Top-most window repository object with Availability tab displayed
'  strProduct = Name of product
'  strLevel = Name of packaging level
'  strFormat = Name of SPT Format
'  strSize = Block size [optional - blank value with ignore block size and disable all rows that match parameters]
'  boolFirstOnly = True to only disable the first row matching strSize; otherwise False to disable all rows matching strSize
' RETURNS: TRUE if found and disabled; otherwise FALSE if not found or disabled
' NOTE: Will log info to result report
Function DisableAvailableList(ByVal objWindow, ByVal strProduct, ByVal strLevel, ByVal strFormat, ByVal strSize, ByVal boolFirstOnly)
	Print Now & " - Disable List with Product=" & strProduct & " Level=" & strLevel & " Format=" & strFormat & " First=" & boolFirstOnly & " Size=" & strSize
	reporter.ReportNote "Disable List with Product=" & strProduct & " Level=" & strLevel & " Format=" & strFormat & " First=" & boolFirstOnly & " Size=" & strSize
	
	DisableAvailableList = False
	Dim objDataGrid 
	Set objDataGrid = objWindow.SwfTable("dtgrdRange")
	
	Dim arrData : arrData = GetDataGridData(objDataGrid, Array("Product Name", "Type Name", "SPT Format Name", "Block Size Received", "Completed"))	
	Dim intRowCount : intRowCount = objDataGrid.RowCount
	Dim row

	For row = 0 To intRowCount-1	' each row
		If (arrData(0,row) = strProduct) And (arrData(1,row) = strLevel) And (arrData(2,row) = strFormat) And (arrData(4,row) = False) Then	' match on primary keys
			If (arrData(3,row) = strSize) Or StringIsNullOrEmpty(strSize) Then	' match on secondary key
				' select and disable
				objDataGrid.SelectRow row
			
				objWindow.SwfButton("btnDisable").Click
				If objWindow.Dialog("dlgConfirmExit").Exist(2) Then
					objWindow.Dialog("dlgConfirmExit").WinButton("btnYes").Click
					reporter.ReportEvent micPass, "Disable List", "Disabled List for Product=" & strProduct & " Level=" & strLevel & " Format=" & strFormat & " Size=" & strSize
					DisableAvailableList = True
					If boolFirstOnly And (Not StringIsNullOrEmpty(strSize)) Then	' found first row matching strSize
						Exit For
					End If
				End If ' dialog
			End If ' match on secondary
		End If ' match on primary
	Next ' row
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Disables a SFDA List of SPT Numbers
'  objWindow = Top-most window repository object with Availability tab displayed
'  strResourceCode = Value of resource code
'  strStart = Start SPT Number of range [optional - blank value with ignore start number and disable all rows that match resource code]
' RETURNS: TRUE if found and disabled; otherwise FALSE if not found or disabled
' NOTE: Will log info to result report
Function DisableAvailableSFDA(ByVal objWindow, ByVal strResourceCode, ByVal strStart)
	Print Now & " - Disable SFDA with Resource=" & strResourceCode & " and Start=" & strStart
	reporter.ReportNote "Disable SFDA with Resource=" & strResourceCode & " and Start=" & strStart
	
	DisableAvailableSFDA = False
	Dim objDataGrid 
	Set objDataGrid = objWindow.SwfTable("dtgrdRange")
	
	Dim arrData : arrData = GetDataGridData(objDataGrid, Array("Resource Code", "Start Number"))	
	Dim intRowCount : intRowCount = objDataGrid.RowCount
	Dim row

	For row = 0 To intRowCount-1	' each row
		If (arrData(0,row) = strResourceCode) Then	' match on primary key
			If (arrData(1,row) = strStart) Or StringIsNullOrEmpty(strStart) Then	' match on secondary key
				' select and disable
				objDataGrid.SelectRow row
			
				objWindow.SwfButton("btnDisable").Click
				If objWindow.Dialog("dlgConfirmExit").Exist(2) Then
					objWindow.Dialog("dlgConfirmExit").WinButton("btnYes").Click
					reporter.ReportEvent micPass, "Disable SFDA", "Disabled SFDA for ResourceCode " & strResourceCode & " with block starting at " & arrData(1,row)
					DisableAvailableSFDA = True
					
					If Not StringIsNullOrEmpty(strStart) Then ' searching for specific row
						Exit For         ' got the row, skip the rest
					End If
				End If ' dialog
			End If ' match on secondary
		End If ' match on primary
	Next ' row
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Delete a preprinted/reel List of SPT Numbers
'  objWindow = Top-most window repository object with Availability tab displayed
'  strReelId = Id of reel to delete
' RETURNS: TRUE if found and disabled; otherwise FALSE if not found or disabled
' NOTE: Will log info to result report
Function DeleteAvailablePreprinted(ByVal objWindow, ByVal strReelId)
	Print Now & " - Delete Preprinted with ReelId=" & strReelId
	reporter.ReportNote "Delete Preprinted with ReelId=" & strReelId
	
	DeleteAvailablePreprinted = False

	' locate and select the reel
	Dim rowCount : rowCount = objWindow.SwfTable("dtgrdReels").RowCount
	Dim row
	For row = 0 To rowCount-1	' each row
		If objWindow.SwfTable("dtgrdReels").GetCellData (row,"Label Reel Number") = strReelId Then
			objWindow.SwfTable("dtgrdReels").ClickCell row,""	' click "X" (first cell of row)
			Exit For
		End If
	Next
	
	If objWindow.Dialog("dlgDelete").Exist(3) Then
		objWindow.Dialog("dlgDelete").WinButton("btnYes").Click
		reporter.ReportEvent micPass, "Delete Preprinted Labels", "Delete Preprinted Labels for ReelId " & strReelId
		DeleteAvailablePreprinted = True
	Else
		objWindow.Dialog("dlgDelete").WinButton("btnOK").Click	' ??? correct ???
		reporter.ReportEvent micFail, "Delete Preprinted Labels", "Failed to delete Preprinted Labels for ReelId " & strReelId
	End If	
End Function


''' ************************************************
''' **** Secondary Helper Functions/Subroutines ****
''' ************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify number of displayed topmenus and submenus (does NOT include submenus not visible due to collapsed topmenu)
'  objMenuTree = Menu tree repository object
'  intExpectedCount = Expected number of displayed menus
' NOTE: Will log info to result report
Sub CheckDisplayedMenuCount(ByVal objMenuTree, ByVal intExpectedCount)
	Dim countVisible : countVisible = objMenuTree.GetItemsCount()

	If countVisible <> intExpectedCount Then
		reporter.ReportEvent micFail, "Menu Count", "Menu has " & countVisible & " item(s) displayed; it was expcted to be " & intExpectedCount & " item(s)"
	Else
		reporter.ReportEvent micPass, "Menu Count", "Menu has " & countVisible & " item(s) displayed"
	End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify if specified screen is displayed by matching titles
'  objLabel = Label repository object of specified screen
'  strTitle = Expected title text 
' RETURN: TRUE if screen title matches; otherwise FALSE
' NOTE: Does not log any info to result report
Function IsScreenDisplayed(ByVal objLabel, ByVal strTitle)
	IsScreenDisplayed = False
	On Error Resume Next
	If Not objLabel.Exist(5) Then
		Exit Function
	End If
	
	IsScreenDisplayed = objLabel.CheckProperty("Text", strTitle, 10)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get ErrorId of latest record in Errors Log
' RETURN: last Error Log ErrorId
' NOTE: Does not log any info to result report
Function LastErrorLogId
	LastErrorLogId = 0
	
	Dim arrData, strId
	ExecuteSQL GetConnectionString, "SELECT MAX(ErrorId) FROM [Guardian].[ErrorsLog]", Null, Null, arrData	
	strId = arrData(0,0)
	
	If Not StringIsNullOrEmpty(strId) Then
		LastErrorLogId = strId
	End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get EventId of latest record in Events Log
' RETURN: last Events Log EventId
' NOTE: Does not log any info to result report
Function LastEventLogId
	LastEventLogId = 0
	
	Dim arrData, strId
	ExecuteSQL GetConnectionString, "SELECT MAX(ErrorId) FROM [Guardian].[EventsLog]", Null, Null, arrData	
	strId = arrData(0,0)
	
	If Not StringIsNullOrEmpty(strId) Then
		LastEventLogId = strId
	End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get Id of latest record in Provision Log
' RETURN: last Provision Log Id
' NOTE: Does not log any info to result report
Function LastProvisionLogId
	LastProvisionLogId = 0
	
	Dim arrData(), arrColumns(), strId
	Dim arrParams(3)
	arrParams(0) = "0"
	arrParams(1) = Null
	arrParams(2) = Null
	arrParams(3) = Null
	ExecuteSQL GetConnectionString, "Guardian.usp_GetMsgQueue", arrParams, Array("Id"), arrData
	If NumberOfDimensions(arrData) = 1  Then
		strId = arrData(UBound(arrData))	
	End If
	
	If Not StringIsNullOrEmpty(strId) Then
		LastProvisionLogId = strId
	End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get Id of latest record in Notification Log
' RETURN: last Notification Log Id
' NOTE: Does not log any info to result report
Function LastNotificationLogId
	LastNotificationLogId = 0
	
	Dim arrData(), arrColumns(), strId
	Dim arrParams(3)
	arrParams(0) = "1"
	arrParams(1) = Null
	arrParams(2) = Null
	arrParams(3) = Null
	ExecuteSQL GetConnectionString, "Guardian.usp_GetMsgQueue", arrParams, Array("Id"), arrData
	If NumberOfDimensions(arrData) = 1  Then
		strId = arrData(UBound(arrData))	
	End If
	
	If Not StringIsNullOrEmpty(strId) Then
		LastNotificationLogId = strId
	End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get specified columns in Events Log starting at a specific reference id
'  arrColumns = Array of column names to retrieve {NULL to get all columns}
'  intId = Starting event id 
' RETURN: Array of contents
' NOTE: Does not log any info to result report
Function GetEventLog(arrColumns, intId)
	Dim arrData
	ExecuteSQL GetConnectionString, "SELECT * FROM [Guardian].[EventsLog] WHERE EventId >= " & intId, Null, arrColumns, arrData	
	GetEventLog = arrData
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get specified columns in Errors Log starting at a specific reference id
'  arrColumns = Array of column names to retrieve {NULL to get all columns}
'  intId = Starting error id 
' RETURN: Array of contents
' NOTE: Does not log any info to result report
Function GetErrorLog(arrColumns, intId)
	Dim arrData
	ExecuteSQL GetConnectionString, "SELECT * FROM [Guardian].[ErrorsLog] WHERE ErrorId >= " & intId, Null, arrColumns, arrData	
	GetErrorLog = arrData
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get specified columns in Provision Log starting at a specific reference id
'  arrColumns = Array of column names to retrieve {NULL to get all columns}
'  intId = Starting provision id
' RETURN: Array of contents
' NOTE: Does not log any info to result report
Function GetProvisionLog(arrColumns, intId)
	Dim arrData
	Dim arrParams(3)
	arrParams(0) = "0"
	arrParams(1) = Null
	arrParams(2) = Null
	arrParams(3) = Null
	ExecuteSQL GetConnectionString, "Guardian.usp_GetMsgQueue", arrParams, arrColumns, arrData
	GetProvisionLog = arrData
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get specified columns in Notification Log starting at a specific reference id
'  arrColumns = Array of column names to retrieve {NULL to get all columns}
'  intId = Starting notification id
' RETURN: Array of contents
' NOTE: Does not log any info to result report
Function GetNotificationLog(arrColumns, intId)
	Dim arrData
	Dim arrParams(3)
	arrParams(0) = "1"
	arrParams(1) = Null
	arrParams(2) = Null
	arrParams(3) = Null
	ExecuteSQL GetConnectionString, "Guardian.usp_GetMsgQueue", arrParams, arrColumns, arrData
	GetNotificationLog = arrData
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' END OF SCRIPT