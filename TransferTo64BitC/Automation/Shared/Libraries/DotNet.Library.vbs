
' HEADER
'------------------------------------------------------------------'
'    Description:  .NET General Library                            '
'                                                                  '
'        Project:  .NET Application Automation                     '
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
'  20170421  v2.2      RNiedzwiecki  added subroutine SetFieldOnScreen
'  20170401  v2.1      RNiedzwiecki  added functions LoadConfigFile, GetConfigNodes
'  20170201  v2.0      RNiedzwiecki  CODE SPLIT from Guardian Library.qfl; this file only contains .NET object functionality
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
' - Configuration file containing path \UFT\Screens\* (e.g. UftConfig.xml)
' - Environment variables are loaded in UFT (File > Settings > Environment > User-defined > Load variables from external file)  [version 1.2+]
' - For legacy scripts using UftConfig.xml  [version 1.x]:
'    - Test script must set configuration file path using SetConfigFile()
'      or
'    - Project shall have a datasheet with the name "Global"
'    - Global datasheet shall have a column with the name "ConfigFile"
'    - ConfigFile column shall have one value containing the path+filename of the XML configuration file (e.g. c:\uft\TestPlan\UftConfig.xml)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' START OF SCRIPT
Option Explicit
Public gblConfig	 ' global link to XML configuration metadata (will be auto-loaded on first access)  [a.k.a. XML configuration file]
Public gblConfigFile ' global configuration file (will be test by parent Test script using SetConfigFile subroutine)

''' ***************************************
''' **** Primary Functions/Subroutines ****
''' ***************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a button field
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the button field repository object
'  isVisible = TRUE if button should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if button should be enabled; otherwise FALSE for disabled
'  strValue = Regex displayed label on the button field (Null=ignore check)
'  strOperator = Match operation to perform on value (e.g. LT|LE|RX|EQ|NE|GE|GT)
'  strTooltip = Tooltip text (Null=ignore check)
' NOTE: Will log info to result report
Sub CheckButtonState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal strValue, ByVal strOperator, ByVal strTooltip)
	Const strStep = "CheckButtonState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfButton(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "Button '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "Button '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "Button '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	objField.CheckProperty "Visible", isVisible
	objField.CheckProperty "Enabled", isEnabled
	
	If Not IsNull(strValue) Then
		VerifyProperty objField, "Text", strOperator, strValue, 1
	End If
	If Not IsNull(strTooltip) And isEnabled Then
		VerifyTooltip objParentWindow, objField, strStep, strTooltip
	End If

	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a label field
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the label field repository object
'  isVisible = TRUE if label should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if label should be enabled; otherwise FALSE for disabled
'  strValue = Regex displayed value of the label field (Null=ignore check)
'  strOperator = Match operation to perform on value (e.g. LT|LE|RX|EQ|NE|GE|GT)
'  strTooltip = Tooltip text (Null=ignore check)
' NOTE: Will log info to result report
Sub CheckLabelState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal strValue, ByVal strOperator, ByVal strTooltip)
	Const strStep = "CheckLabelState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfLabel(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "Label '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "Label '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "Label '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	objField.CheckProperty "Visible", isVisible
	objField.CheckProperty "Enabled", isEnabled
	
	If Not IsNull(strValue) Then
		VerifyProperty objField, "Text", strOperator, strValue, 1
	End If
	If Not IsNull(strTooltip) And isEnabled Then
		VerifyTooltip objParentWindow, objField, strStep, strTooltip
	End If

	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a groupbox field
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the groupbox field repository object
'  isVisible = TRUE if groupbox should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if groupbox should be enabled; otherwise FALSE for disabled
'  strValue = Regex displayed text of the groupbox field (Null=ignore check)
'  strOperator = Match operation to perform on value (e.g. LT|LE|RX|EQ|NE|GE|GT)
'  strTooltip = Tooltip text (Null=ignore check)
' NOTE: Will log info to result report
Sub CheckGroupBoxState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal strValue, ByVal strOperator, ByVal strTooltip)
	Const strStep = "CheckGroupBoxState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfObject(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "Groupbox '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "Groupbox '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "Groupbox '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	objField.CheckProperty "Visible", isVisible
	objField.CheckProperty "Enabled", isEnabled
	
	If Not IsNull(strValue) Then
		VerifyProperty objField, "Text", strOperator, strValue, 1
	End If
	If Not IsNull(strTooltip) And isEnabled Then
		VerifyTooltip objParentWindow, objField, strStep, strTooltip
	End If
	
	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a textbox field
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the textbox field repository object
'  isVisible = TRUE if field should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if field should be enabled; otherwise FALSE for disabled
'  isReadOnly = TRUE if field should be read-only; otherwise FALSE for full edit read-write
'  strValue = Regex default value of the textbox field (Null=ignore check)
'  strOperator = Match operation to perform on value (e.g. LT|LE|RX|EQ|NE|GE|GT)
'  strTooltip = Tooltip regex text (Null=ignore check)
'  maxLength = Maximum length of field (Null=ignore check)
' NOTE: Will log info to result report
Sub CheckTextBoxState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal isReadOnly, ByVal strValue, ByVal strOperator, ByVal strTooltip, ByVal maxLength)
	Const strStep = "CheckTextBoxState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfEdit(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "Textbox '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "Textbox '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "Textbox '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	objField.CheckProperty "Visible", isVisible
	objField.CheckProperty "Enabled", isEnabled
	objField.CheckProperty "ReadOnly", isReadOnly
	
	If Not IsNull(maxLength) Then
		objField.CheckProperty "MaxLength", maxLength
	End If
	If Not IsNull(strValue) Then
		VerifyProperty objField, "Text", strOperator, strValue, 1
	End If
	If Not IsNull(strTooltip) And isEnabled Then
		VerifyTooltip objParentWindow, objField, strStep, strTooltip
	End If
	
	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a multi-line textbox field
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the textbox field repository object
'  isVisible = TRUE if field should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if field should be enabled; otherwise FALSE for disabled
'  isReadOnly = TRUE if field should be read-only; otherwise FALSE for full edit read-write
'  strValue = Regex default text of the textbox field (Null=ignore check)
'  strOperator = Match operation to perform on value (e.g. LT|LE|RX|EQ|NE|GE|GT)
'  strTooltip = Tooltip text (Null=ignore check)
'  maxLength = Maximum length of field (Null=ignore check)
' NOTE: Will log info to result report
Sub CheckMultiTextBoxState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal isReadOnly, ByVal strValue, ByVal strOperator, ByVal strTooltip, ByVal maxLength)
	Const strStep = "CheckMultilineTextBoxState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfEditor(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "Multiline Textbox '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "Multiline Textbox '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "Multiline Textbox '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	objField.CheckProperty "Visible", isVisible
	objField.CheckProperty "Enabled", isEnabled
	objField.CheckProperty "ReadOnly", isReadOnly
	If Not IsNull(maxLength) Then
		objField.CheckProperty "MaxLength", maxLength
	End If
	If Not IsNull(strValue) Then
		VerifyProperty objField, "Text", strOperator, strValue, 1
	End If
	If Not IsNull(strTooltip) And isEnabled Then
		VerifyTooltip objParentWindow, objField, strStep, strTooltip
	End If
	
	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a dropdown edit field
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the dropdown edit field repository object
'  strValue = Regex default value of the textbox field (Null=ignore check)
'  strOperator = Match operation to perform on value (e.g. LT|LE|RX|EQ|NE|GE|GT)
'  strTooltip = Tooltip regex text (Null=ignore check)
'  maxLength = Maximum length of field (Null=ignore check)
' NOTE: Will log info to result report
Sub CheckDropDownEditState( ByVal objParentWindow, ByVal strFieldName, ByVal strValue, ByVal strOperator, ByVal strTooltip, ByVal maxLength)
	Const strStep = "CheckDropDownEditState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.WinEdit(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	If Not IsNull(maxLength) Then
		objField.CheckProperty "MaxLength", maxLength
	End If
	If Not IsNull(strValue) Then
		VerifyProperty objField, "Text", strOperator, strValue, 1
	End If
	If Not IsNull(strTooltip) Then
		VerifyTooltip objParentWindow, objField, strStep, strTooltip
	End If
	
	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a dropdown field
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the dropdown field repository object
'  isVisible = TRUE if dropdown should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if dropdown should be enabled; otherwise FALSE for disabled
'  arrChoices = List of available dropdown choices (Null=ignore check)
'  isPartialList = TRUE if arrChoices is a partial list of choices; otherwise FALSE for full list
'  strValue = Regex default value of the dropdown field (Null=ignore check)
'  strOperator = Match operation to perform on value (e.g. LT|LE|RX|EQ|NE|GE|GT)
'  strTooltip = Tooltip text (Null=ignore check)
' NOTE: Will log info to result report
Sub CheckDropDownState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal arrChoices, ByVal isPartialList, ByVal strValue, ByVal strOperator, ByVal strTooltip)
	Const strStep = "CheckDropDownState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfComboBox(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "Dropdown '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "Dropdown '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "Dropdown '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	Dim countShown, countExpected, strItem, i
	
	objField.CheckProperty "Visible", isVisible
	objField.CheckProperty "Enabled", isEnabled

	' verify available choices match
	If Not IsNull(arrChoices) AND IsArray(arrChoices) AND NumberOfDimensions(arrChoices) = 1 Then	
		countShown = objField.GetItemsCount()
		countExpected = Ubound(arrChoices) - lbound(arrChoices) + 1
		
		If countExpected > 0 Then
			If Not IsNull(isPartialList) Then
				If isPartialList Then	' verify partial list provided
					reporter.ReportEvent micPass, "Dropdown Count", strFieldName & " has " & countShown & " item(s)"
					If countExpected < countShown Then
						countShown = countExpected		' limit verification to partial list expected (not visible list)
					End If
					For i = 0 To countShown - 1
						strItem = objField.GetItem(i)
						
						If strItem = arrChoices(i) Then
							reporter.ReportEvent micPass, "Dropdown List", strFieldName & " item(" & i & ") has the actual value '" & strItem & "'"		
						Else
							reporter.ReportEvent micFail, "Dropdown List", strFieldName & " item(" & i & ") has the actual value '" & strItem & "'; it was expected to be '" & arrChoices(i) & "'"
						End If
					Next					
				End If
			Else	' verify all visible choices
				If countShown <> countExpected Then
					reporter.ReportEvent micFail, "Dropdown Count", strFieldName & " has " & countShown & " item(s); it was expected to have " & countExpected & " item(s)"
				Else
					reporter.ReportEvent micPass, "Dropdown Count", strFieldName & " has " & countShown & " item(s)"
					For i = 0 To countShown - 1
						strItem = objField.GetItem(i)
						
						If strItem = arrChoices(i) Then
							reporter.ReportEvent micPass, "Dropdown List", strFieldName & " item(" & i & ") has the actual value '" & strItem & "'"		
						Else
							reporter.ReportEvent micFail, "Dropdown List", strFieldName & " item(" & i & ") has the actual value '" & strItem & "'; it was expected to be '" & arrChoices(i) & "'"
						End If
					Next
				End If
			End If		
		End If
	End If

	If Not IsNull(strValue) Then	
		VerifyProperty objField, "Text", strOperator, strValue, 1
	End If
	If Not IsNull(strTooltip) And isEnabled Then
		VerifyTooltip objParentWindow, objField, strStep, strTooltip
	End If

	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a listbox field
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the listbox field repository object
'  isVisible = TRUE if listbox should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if listbox should be enabled; otherwise FALSE for disabled
'  arrChoices = List of available listbox choices (Null=ignore check)
'  strTooltip = Tooltip text (Null=ignore check)
' NOTE: Will log info to result report
Sub CheckListBoxState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal arrChoices, ByVal strTooltip)
	Const strStep = "CheckListBoxState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfList(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "ListBox '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "ListBox '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "ListBox '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	Dim countShown, countExpected, strItem, i
	
	objField.CheckProperty "Visible", isVisible
	objField.CheckProperty "Enabled", isEnabled

	' verify available choices match
	If Not IsNull(arrChoices) AND IsArray(arrChoices) AND NumberOfDimensions(arrChoices) = 1 Then	
		countShown = objField.GetItemsCount()
		countExpected = Ubound(arrChoices) - lbound(arrChoices) + 1
		
		If countExpected > 0 Then
			If countShown <> countExpected Then
				reporter.ReportEvent micFail, "ListBox Count", "ListBox list has " & countShown & " item(s); it was expected to have " & countExpected & " item(s)"
			Else
				reporter.ReportEvent micPass, "ListBox Count", "ListBox list has " & countShown & " item(s)"
				For i = 0 To countShown - 1
					strItem = objField.GetItem(i)
					
					If strItem = arrChoices(i) Then
						reporter.ReportEvent micPass, "ListBox List", "Item(" & i & ") has the actual value '" & strItem & "'"		
					Else
						reporter.ReportEvent micFail, "ListBox List", "Item(" & i & ") has the actual value '" & strItem & "'; it was expected to be '" & arrChoices(i) & "'"
					End If		
				Next	
			End If		
		End If
	End If

	If Not IsNull(strTooltip) And isEnabled Then
		VerifyTooltip objParentWindow, objField, strStep, strTooltip
	End If

	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a checkbox field
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the checkbox field repository object
'  isVisible = TRUE if checkbox should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if checkbox should be enabled; otherwise FALSE for disabled
'  strValue = Regex displayed label on the checkbox field (Null=ignore check)
'  strOperator = Match operation to perform on value (e.g. LT|LE|RX|EQ|NE|GE|GT)
'  isChecked = Default value of the checkbox field (Null=ignore check)
'  strTooltip = Tooltip text (Null=ignore check)
' NOTE: Will log info to result report
Sub CheckCheckBoxState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal strValue, ByVal strOperator, ByVal isChecked, ByVal strTooltip)
	Const strStep = "CheckCheckBoxState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfCheckBox(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "CheckBox '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "CheckBox '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "CheckBox '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	objField.CheckProperty "Visible", isVisible
	objField.CheckProperty "Enabled", isEnabled
	objField.CheckProperty "Checked", isChecked
	
	If Not IsNull(strValue) Then
		VerifyProperty objField, "Text", strOperator, strValue, 1
	End If
	If Not IsNull(strTooltip) And isEnabled Then
		VerifyTooltip objParentWindow, objField, strStep, strTooltip
	End If

	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a radio button field
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the radio button field repository object
'  isVisible = TRUE if radio button should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if radio button should be enabled; otherwise FALSE for disabled
'  strValue = Regex displayed label on the radio button field (Null=ignore check)
'  strOperator = Match operation to perform on value (e.g. LT|LE|RX|EQ|NE|GE|GT)
'  isChecked = Default value of the radio button field (Null=ignore check)
'  strTooltip = Tooltip text (Null=ignore check)
' NOTE: Will log info to result report
Sub CheckRadioButtonState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal strValue, ByVal strOperator, ByVal isChecked, ByVal strTooltip)
	Const strStep = "CheckRadioButtonState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfRadioButton(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "RadioButton '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "RadioButton '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "RadioButton '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	objField.CheckProperty "Visible", isVisible
	objField.CheckProperty "Enabled", isEnabled
	objField.CheckProperty "Checked", isChecked

	If Not IsNull(strValue) Then
		VerifyProperty objField, "Text", strOperator, strValue, 1
	End If
	If Not IsNull(strTooltip) And isEnabled Then
		VerifyTooltip objParentWindow, objField, strStep, strTooltip
	End If

	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a numeric spinner field
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the numeric field repository object
'  isVisible = TRUE if field should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if field should be enabled; otherwise FALSE for disabled
'  intValue = Default value of the field
'  strOperator = Match operation to perform on value (e.g. LT|LE|RX|EQ|NE|GE|GT)
'  intMin = Minimum value of the field
'  lngMax = Maximum value of the field
'  intStep = Incremental step of the field
'  strTooltip = Tooltip text (Null=ignore check)
' NOTE: Will log info to result report
Sub CheckNumberSpinnerState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal intValue, ByVal strOperator, ByVal intMin, ByVal lngMax, ByVal intStep, ByVal strTooltip)
	Const strStep = "CheckNumberSpinnerState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfSpin(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "NumberSpinner '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "NumberSpinner '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "NumberSpinner '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	objField.CheckProperty "Visible", isVisible
	objField.CheckProperty "Enabled", isEnabled
	objField.CheckProperty "Minimum", intMin
	objField.CheckProperty "Maximum", lngMax
	objField.CheckProperty "Increment", intStep
	VerifyProperty objField, "Value", strOperator, intValue, 1
	If Not IsNull(strTooltip) And isEnabled Then
		VerifyTooltip objParentWindow, objField, strStep, strTooltip
	End If

	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a calendar field
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the calendar field repository object
'  isVisible = TRUE if field should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if field should be enabled; otherwise FALSE for disabled
'  strValue = Default value of the calendar field (Null=ignore check)
'  strOperator = Match operation to perform on value (e.g. LT|LE|RX|EQ|NE|GE|GT)
'  strTooltip = Tooltip text (Null=ignore check)
' NOTE: Will log info to result report
Sub CheckCalendarState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal strValue, ByVal strOperator, ByVal strTooltip)
	Const strStep = "CheckCalendarState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfCalendar(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "Calendar '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "Calendar '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "Calendar '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	objField.CheckProperty "Visible", isVisible
	objField.CheckProperty "Enabled", isEnabled

	If Not IsNull(strValue) Then
		If StringIsNullOrEmpty(strValue) Then
			strValue = Date
		End If
		VerifyProperty objField, "Date", strOperator, strValue, 1
	End If
	If Not IsNull(strTooltip) And isEnabled Then
		VerifyTooltip objParentWindow, objField, strStep, strTooltip
	End If
	
	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a datagrid
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the datagrid repository object
'  isVisible = TRUE if datagrid should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if datagrid should be enabled; otherwise FALSE for disabled
'  arrColumns = List of column names expected in datagrid
'  arrExpectedData = Array of values expected in datagrid
' NOTE: Will log info to result report
Sub CheckDataGridState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal arrColumns, ByVal arrExpectedData)
	Const strStep = "CheckDataGridState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfTable(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "DataGrid '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "DataGrid '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "DataGrid '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	objField.CheckProperty "Visible", isVisible
'	If (UBound(arrExpectedData,1) = 0) Then
'		isEnabled = false
'	End If
'	objField.CheckProperty "Enabled", isEnabled

	Dim arrVisibleColumns()
	Dim arrGridData : arrGridData = GetDataGridContents( objField, false, arrVisibleColumns)

	If UBound(arrColumns) > 0 Then
		CheckIfArraysMatch arrColumns, arrVisibleColumns, false, "ColumnHeaders"	' verify columns match		
	End If
	If Not IsNull(arrExpectedData) Then
		CheckIfArraysMatch arrExpectedData, arrGridData, false, "GridData"			' verify grid data matches		
	End If

	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a treeview
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the treeview repository object
'  isVisible = TRUE if treeview should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if treeview should be enabled; otherwise FALSE for disabled
'  arrNodes = List of visible nodes expected in datagrid
' NOTE: Will log info to result report
Sub CheckTreeViewState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal arrNodes)
	Const strStep = "CheckTreeViewState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfTreeView(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "TreeView '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "TreeView '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "TreeView '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	objField.CheckProperty "Visible", isVisible
	objField.CheckProperty "Enabled", isEnabled

	Dim intCountVisible : intCountVisible = objField.GetItemsCount()
	Dim intExpectedCount : intExpectedCount = Ubound(arrNodes) - lbound(arrNodes) + 1

	If intCountVisible <> intExpectedCount Then
		reporter.ReportEvent micFail, strStep, "Tree has " & intCountVisible & " item(s) displayed; it was expcted to be " & intExpectedCount & " item(s)"
	Else
		reporter.ReportEvent micPass, strStep, "Tree has " & intCountVisible & " item(s) displayed"
	End If

	Dim strItem, i
	For i = 0 To intCountVisible - 1
		strItem = objField.GetItem(i)
		
		If strItem = arrNodes(i) Then
			reporter.ReportEvent micPass, strStep, "Item " & i & " has the actual value '" & strItem & "'"
		Else
			reporter.ReportEvent micFail, strStep, "Item " & i & " has the actual value '" & strItem & "'; it was expected to be '" & arrNodes(i) & "'"
		End If		
	Next

	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a tab group
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the tab group repository object
'  isVisible = TRUE if tab group should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if tab group should be enabled; otherwise FALSE for disabled
'  intTabCount = Numbers of tabs expected within the group
' NOTE: Will log info to result report
Sub CheckTabGroupState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal intTabCount)
	Const strStep = "CheckTabGroupState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfTab(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "TabGroup '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "TabGroup '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "TabGroup '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	objField.CheckProperty "Visible", isVisible
	objField.CheckProperty "Enabled", isEnabled
	objField.CheckProperty "TabCount", intTabCount

	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the properties of a tab page
'  objParentWindow = Parent window repository object
'  strFieldName = Name of the tab screen repository object
'  isVisible = TRUE if the tab should be visible; otherwise FALSE for non-visible
'  isEnabled = TRUE if the tab should be enabled; otherwise FALSE for disabled
'  strValue = Regex display label of the tab
'  strOperator = Match operation to perform on value (e.g. LT|LE|RX|EQ|NE|GE|GT)
' NOTE: Will log info to result report
Sub CheckTabState( ByVal objParentWindow, ByVal strFieldName, ByVal isVisible, ByVal isEnabled, ByVal strValue, ByVal strOperator)
	Const strStep = "CheckTabState"
	print Now & " " & strStep & " " & strFieldName
	Dim objField 
	Set objField = objParentWindow.SwfObject(strFieldName)
	Dim isExist : isExist = objField.Exist(0)

	' verify visibility
	If Not isVisible And Not isExist Then
		reporter.ReportEvent micPass, strStep, "Tab '" & strFieldName & "' is not visible, as expected"
		Exit Sub
	End If
	If Not isVisible And isExist Then
		reporter.ReportEvent micFail, strStep, "Tab '" & strFieldName & "' is visible and it was expected to be not visible"
		Exit Sub
	End If
	If isVisible And Not isExist Then
		reporter.ReportEvent micFail, strStep, "Tab '" & strFieldName & "' is not visible and it was expected to be visible"
		Exit Sub
	End If

	objField.CheckProperty "Visible", isVisible
	objField.CheckProperty "Enabled", isEnabled
	VerifyProperty objField, "Text", strOperator, strValue, 1

	Set objField = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the state of the entire screen (checks the properties of each field/control)
'  objParentWindow = Parent window repository object
'  strScreenName = Name of the screenshot in the global configuration file
' NOTE: Will log info to result report
Sub CheckScreenState( ByVal objParentWindow, ByVal strScreenName )
	Const strStep = "CheckScreenState"
	print Now & " " & strStep & " " & strScreenName
	Dim xmlRoot, colNodes, objNode, strName, isVisible, isEnabled, strValue, strOperator, strTooltip, i

	If Not objParentWindow.Exist(0) Then
		reporter.ReportEvent micFail, strStep, objParentWindow.ToString & " not found; aborting " & strStep
		Exit Sub
	End If

	' get root node for screen
	Set xmlRoot = GetConfigNode("/UFT/Screens/" & strScreenName)
	If xmlRoot Is Nothing Then
		reporter.ReportEvent micFail, strStep, "Screen name " & strScreenName & " not found in configuration file.  Aborting process..."
		Exit Sub
	End If

   	strValue = xmlRoot.getAttribute("value")					' optional
   	strOperator = xmlRoot.getAttribute("operator")				' optional
   	If Not StringIsNullOrEmpty(strValue) Then
		VerifyProperty objParentWindow, "Text", strOperator, strValue, 1
   	End If

	'''''' verify tab groups
	Set colNodes = xmlRoot.selectNodes("TabGroup")	' get list of tab groups from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	
		' get list of tabs
		Dim colTabs, objTab, arrTabs()
		Set colTabs = objNode.selectNodes("Tab")
		i = 0
		
		For Each objTab in colTabs
			ReDim Preserve arrTabs(i)
			arrTabs(i) = objTab.getAttribute("name")
		 	i = i + 1
		Next

		Set colTabs = Nothing
		CheckTabGroupState objParentWindow, strName, isVisible, isEnabled, UBound(arrTabs)+1
		ReDim arrTabs(-1)
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - TabGroups", "TabGroup check completed for screen " & strScreenName

	'''''' verify tabs
	Set colNodes = xmlRoot.selectNodes("TabGroup/Tab")	' get list of tabs from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	strValue = objNode.getAttribute("value")
	   	strOperator = objNode.getAttribute("operator")				' optional
	   	
		CheckTabState objParentWindow, strName, isVisible, isEnabled, strValue, strOperator
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - Tabs", "Tab check completed for screen " & strScreenName

	'''''' verify buttons
	Set colNodes = xmlRoot.selectNodes("Button")		' get list of buttons from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	strValue = objNode.getAttribute("value")
	   	strOperator = objNode.getAttribute("operator")				' optional
	   	strTooltip = objNode.getAttribute("tooltip")				' optional
		CheckButtonState objParentWindow, strName, isVisible, isEnabled, strValue, strOperator, strTooltip
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - Buttons", "Button check completed for screen " & strScreenName
	
	'''''' verify labels
	Set colNodes = xmlRoot.selectNodes("Label")			' get list of labels from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	strValue = objNode.getAttribute("value")
	   	strOperator = objNode.getAttribute("operator")				' optional
	   	strTooltip = objNode.getAttribute("tooltip")				' optional
		CheckLabelState objParentWindow, strName, isVisible, isEnabled, strValue, strOperator, strTooltip
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - Labels", "Label check completed for screen " & strScreenName
	
	'''''' verify groupboxes
	Set colNodes = xmlRoot.selectNodes("GroupBox")		' get list of groupboxes from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	strValue = objNode.getAttribute("value")
	   	strOperator = objNode.getAttribute("operator")				' optional
	   	strTooltip = objNode.getAttribute("tooltip")				' optional
		CheckGroupBoxState objParentWindow, strName, isVisible, isEnabled, strValue, strOperator, strTooltip
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - GroupBoxes", "GroupBox check completed for screen " & strScreenName
	
	'''''' verify textboxes
	Dim intLength, isReadOnly
	Set colNodes = xmlRoot.selectNodes("TextBox")		' get list of textbox from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	isReadOnly = CBool(objNode.getAttribute("readonly"))
	   	strValue = objNode.getAttribute("value")					' optional
	   	strOperator = objNode.getAttribute("operator")				' optional
	   	strTooltip = objNode.getAttribute("tooltip")				' optional
	   	intLength = CInteger(objNode.getAttribute("maxlength"))		' optional
		CheckTextBoxState objParentWindow, strName, isVisible, isEnabled, isReadOnly, strValue, strOperator, strTooltip, intLength
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - TextBoxes", "TextBox check completed for screen " & strScreenName
	
	'''''' verify multi-line textboxes
	Set colNodes = xmlRoot.selectNodes("MultilineTextBox")		' get list of multiline textbox from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	isReadOnly = CBool(objNode.getAttribute("readonly"))
	   	strValue = objNode.getAttribute("value")					' optional
	   	strOperator = objNode.getAttribute("operator")				' optional
	   	strTooltip = objNode.getAttribute("tooltip")				' optional
	   	intLength = CInteger(objNode.getAttribute("maxlength"))		' optional
		CheckMultiTextBoxState objParentWindow, strName, isVisible, isEnabled, isReadOnly, strValue, strOperator, strTooltip, intLength
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - Multiline TextBoxes", "Multiline TextBox check completed for screen " & strScreenName
	
	'''''' verify checkboxes
	Dim isChecked
	Set colNodes = xmlRoot.selectNodes("CheckBox")		' get list of checkboxes from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	isChecked = CBool(objNode.getAttribute("checked"))
	   	strValue = objNode.getAttribute("value")					' optional
	   	strOperator = objNode.getAttribute("operator")				' optional
	   	strTooltip = objNode.getAttribute("tooltip")				' optional
		CheckCheckBoxState objParentWindow, strName, isVisible, isEnabled, strValue, strOperator, isChecked, strTooltip
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - CheckBoxes", "CheckBox check completed for screen " & strScreenName
	
	'''''' verify radio buttons
	Set colNodes = xmlRoot.selectNodes("RadioButton")		' get list of radio buttons from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	isChecked = CBool(objNode.getAttribute("checked"))
	   	strValue = objNode.getAttribute("value")					' optional
	   	strOperator = objNode.getAttribute("operator")				' optional
	   	strTooltip = objNode.getAttribute("tooltip")				' optional
		CheckRadioButtonState objParentWindow, strName, isVisible, isEnabled, strValue, strOperator, isChecked, strTooltip
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - RadioButtons", "RadioButton check completed for screen " & strScreenName
	
	'''''' verify number spinners
	Dim intValue, intMin, lngMax, intStep
	Set colNodes = xmlRoot.selectNodes("NumberSpinner")	' get list of numerical spinners from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	intValue = objNode.getAttribute("value")
	   	strOperator = objNode.getAttribute("operator")
	   	intMin = objNode.getAttribute("minimum")
	   	lngMax = objNode.getAttribute("maximum")
	   	intStep = objNode.getAttribute("increment")
	   	strTooltip = objNode.getAttribute("tooltip")				' optional
		CheckNumberSpinnerState objParentWindow, strName, isVisible, isEnabled, intValue, strOperator, intMin, lngMax, intStep, strTooltip
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - NumberSpinners", "NumberSpinner check completed for screen " & strScreenName

	'''''' verify calendars
	Set colNodes = xmlRoot.selectNodes("Calendar")		' get list of calendars from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	strValue = objNode.getAttribute("value")					' optional
	   	strOperator = objNode.getAttribute("operator")				' optional
	   	strTooltip = objNode.getAttribute("tooltip")				' optional
		CheckCalendarState objParentWindow, strName, isVisible, isEnabled, strValue, strOperator, strTooltip
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - Calendars", "Calendar check completed for screen " & strScreenName
	
	'''''' verify dropdowns
	Dim arrChoices(), isPartialList
	Dim colSubNodes, objSubNode
	Set colNodes = xmlRoot.selectNodes("DropDown")		' get list of dropdowns from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	strValue = objNode.getAttribute("value")
	   	strOperator = objNode.getAttribute("operator")				' optional
	   	strTooltip = objNode.getAttribute("tooltip")				' optional
	   	isPartialList = objNode.getAttribute("partiallist")			' optional
	
		' get list of predefined choices (optional)
		Set colSubNodes = objNode.selectNodes("Choice")
		i = 0
		
		For Each objSubNode in colSubNodes
			ReDim Preserve arrChoices(i)
			arrChoices(i) = objSubNode.Text
		 	i = i + 1
		Next
		
		If i = 0 Then		
			' no choices were provided, try a stored procedure to get list of choices (optional)
			Set objSubNode = objNode.selectSingleNode("Procedure")
			If Not objSubNode is Nothing Then
				ExecProcedure_GetArray objSubNode, arrChoices		
			End If		
		End If

		Set objSubNode = Nothing
		Set colSubNodes = Nothing
		CheckDropDownState objParentWindow, strName, isVisible, isEnabled, arrChoices, isPartialList, strValue, strOperator, strTooltip
		ReDim arrChoices(-1)
		
		If isVisible And isEnabled Then
			Set objSubNode = objNode.selectSingleNode("DropDownEdit")
			If Not objSubNode Is Nothing Then								' combobox mode (alternate edit field is available)
			   	strName = objSubNode.getAttribute("name")
			   	strValue = objSubNode.getAttribute("value")
			   	strOperator = objSubNode.getAttribute("operator")			' optional		
			   	intLength = CInteger(objSubNode.getAttribute("maxlength"))	' optional
			   	strTooltip = objSubNode.getAttribute("tooltip")				' optional		
			   	
				CheckDropDownEditState objParentWindow, strName, strValue, strOperator, strTooltip, intLength
				Set objSubNode = Nothing
			End If			
		End If		
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - Dropdowns", "DropDown check completed for screen " & strScreenName

	'''''' verify listboxes
	ReDim arrChoices(-1)
	Set colNodes = xmlRoot.selectNodes("ListBox")		' get list of listboxes from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	'strValue = objNode.getAttribute("value")
	   	strTooltip = objNode.getAttribute("tooltip")				' optional
	
		' get list of predefined choices (optional)
		Set colSubNodes = objNode.selectNodes("Choice")
		i = 0
		
		For Each objSubNode in colSubNodes
			ReDim Preserve arrChoices(i)
			arrChoices(i) = objSubNode.Text
		 	i = i + 1
		Next
		
		If i = 0 Then		
			' no choices were provided, try a stored procedure to get list of choices (optional)
			Set objSubNode = objNode.selectSingleNode("Procedure")
			If Not objSubNode is Nothing Then
				ExecProcedure_GetArray objSubNode, arrChoices		
			End If		
		End If

		Set objSubNode = Nothing
		Set colSubNodes = Nothing
		CheckListBoxState objParentWindow, strName, isVisible, isEnabled, arrChoices, strTooltip
		ReDim arrChoices(-1)
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - ListBoxes", "ListBox check completed for screen " & strScreenName

	'''''' verify data grids
	Dim arrColumns(), arrDbData()
	Set colNodes = xmlRoot.selectNodes("DataGrid")	' get list of data grids from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	
		' get list of columns
		Dim colColumns, objColumn
		Set colColumns = objNode.selectNodes("Column")
		ReDim arrColumns(-1)
		i = 0
		
		For Each objColumn in colColumns
			ReDim Preserve arrColumns(i)
			arrColumns(i) = objColumn.Text
		 	i = i + 1
		Next

		' get procedure for DB data
		Set objSubNode = objNode.selectSingleNode("Procedure")
		If (Not objSubNode is Nothing) And isEnabled Then
			ExecProcedure_GetArray objSubNode, arrDbData
		Else
			ReDim arrDbData(-1,-1)
		End If		

		If Not objSubNode is Nothing Then
			CheckDataGridState objParentWindow, strName, isVisible, isEnabled, arrColumns, arrDbData
		Else			
			CheckDataGridState objParentWindow, strName, isVisible, isEnabled, arrColumns, Null
		End If
		
		Set objSubNode = Nothing
		Set colColumns = Nothing
		ReDim arrColumns(-1)	' clear
		ReDim arrDbData(-1,-1)	' clear
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - DataGrids", "DataGrid check completed for screen " & strScreenName

	'''''' verify treeviews
	Dim arrItems()
	Set colNodes = xmlRoot.selectNodes("TreeView")	' get list of treeviews from config file
	
	For Each objNode in colNodes
	   	strName = objNode.getAttribute("name")
	   	isVisible = CBool(objNode.getAttribute("visible"))
	   	isEnabled = CBool(objNode.getAttribute("enabled"))
	   	
		' get list of nodes
		Dim colItems, objItem
		Set colItems = objNode.selectNodes("Node")
		ReDim arrItems(-1)
		i = 0
		
		For Each objItem in colItems
			ReDim Preserve arrItems(i)
			arrItems(i) = objItem.Text
		 	i = i + 1
		Next

		CheckTreeViewState objParentWindow, strName, isVisible, isEnabled, arrItems
		
		Set colItems = Nothing
		ReDim arrItems(-1)	' clear
	Next
	Set colNodes = Nothing
	reporter.ReportEvent micDone, strStep & " - TreeViews", "TreeView check completed for screen " & strScreenName

	Set xmlRoot = Nothing
	reporter.ReportEvent micDone, strStep, "Completed scanning screen " & strScreenName
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify a specific property matches a specified value
'  objField = Field repository object
'  strProperty = Name of property to be verified
'  strOperator = Operation to perform on value
'  strValue = Value to be verified
'  intTimeout = optional timeout value (Null=ignore)
' NOTE: Will log info to result report
Sub VerifyProperty(ByVal objField, ByVal strProperty, ByVal strOperator, ByVal strValue, ByVal intTimeout)
	If IsNull(strOperator) Then
		strOperator = "EQ"
	End If

	Select Case strOperator
		Case "LT"
			If Not IsNull(intTimeout) Then
				objField.CheckProperty strProperty, micLessThan(strValue), intTimeout
			Else
				objField.CheckProperty strProperty, micLessThan(strValue)
			End If
		Case "LE"
			If Not IsNull(intTimeout) Then
				objField.CheckProperty strProperty, micLessThanOrEqual(strValue), intTimeout
			Else
				objField.CheckProperty strProperty, micLessThanOrEqual(strValue)
			End If
		Case "RX"
			If Not IsNull(intTimeout) Then
				objField.CheckProperty strProperty, micRegExpMatch(strValue), intTimeout
			Else
				objField.CheckProperty strProperty, micRegExpMatch(strValue)
			End If
		Case "EQ"
			If Not IsNull(intTimeout) Then
				objField.CheckProperty strProperty, strValue, intTimeout
			Else
				objField.CheckProperty strProperty, strValue
			End If
		Case "NE"
			If Not IsNull(intTimeout) Then
				objField.CheckProperty strProperty, micNotEqual(strValue), intTimeout
			Else
				objField.CheckProperty strProperty, micNotEqual(strValue)
			End If
		Case "GE"	
			If Not IsNull(intTimeout) Then
				objField.CheckProperty strProperty, micGreaterThanOrEqual(strValue), intTimeout
			Else
				objField.CheckProperty strProperty, micGreaterThanOrEqual(strValue)
			End If
		Case "GT"
			If Not IsNull(intTimeout) Then
				objField.CheckProperty strProperty, micGreaterThan(strValue), intTimeout
			Else
				objField.CheckProperty strProperty, micGreaterThan(strValue)
			End If
		Case Else
			MsgBox "Unknown operator: " + strOperator
	End Select	
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify the tooltip for a specified field/control
'  objParent = Parent window/dialog repository object that contains the field being tested (full parent path)
'  objField = Field object being tested for tooltip
'  strTestName = Name of test step initiating this verification (this will be used on the result report)
'  strExpectedTooltip = Expected tooltip text
' NOTE: Will log info to result report
Sub VerifyTooltip(ByVal objParent, ByVal objField, ByVal strTestName, ByVal strExpectedTooltip)
	objField.MouseMove 10, 10		' trigger display of tooltip 
	wait 2

	If Not objParent.SwfObject("swfname path:=" & objParent.GetTOProperty("swfname")).Exist(5) Then
		reporter.ReportEvent micFail, strTestName,  objField.ToString & " tooltip was not visible"
		Exit Sub		
	End If
	
	' get tooltip text
	Dim strDisplayedTip : strDisplayedTip = objParent.SwfObject("swfname path:=" & objParent.GetTOProperty("swfname")).GetROProperty("text")
	
	If RegExpTest(strExpectedTooltip, strDisplayedTip, false) Then
		reporter.ReportEvent micPass, strTestName, objField.ToString & " tooltip matches regular expression"
	Else
		reporter.ReportEvent micFail, strTestName, objField.ToString & " actual tooltip is '" & strDisplayedTip & "'. The regular expression is '" & strExpectedTooltip & "'"
	End If
	objField.MouseMove -1, -1		' move off field
End Sub


''' **************************************
''' ***** Configuration File access ******
''' **************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Sets the path of the global configuration file
'  strValue = path and filename of configuration file to be used
' NOTE: Does not log any info to result report
Sub SetConfigFile(strValue)
	If IsFileExists(strValue) Then
		gblConfigFile = strValue
	End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Loads global configuration XML file (filename found in GlobalSheet or Environment Variable)
' RETURNS: Collection of XmlNodes or NULL
' NOTE: Does not log any info to result report
Sub LoadConfigFile()
	' auto-load primary global configuration file
	If IsEmpty(gblConfig) or IsNull(gblConfig) Then
		' get configuration file path
		If StringIsNullOrEmpty(gblConfigFile) Then					' check if file was provided by parent Test script
			On Error Resume Next
			SetConfigFile datatable("ConfigFile", dtGlobalSheet)	' check for file in datasheet (backward compatibility)
		End If
		If StringIsNullOrEmpty(gblConfigFile) Then
			On Error Resume Next
			SetConfigFile Environment("App_Metadata")				' check for file in environment variables 
		End If
		If StringIsNullOrEmpty(gblConfigFile) Then
			MsgBox "Configuration file not defined",, "XML Load Error" 
			Exit Sub												' out of luck - abort
		End If
	
		'Set gblConfig = XMLUtil.CreateXMLFromFile(gblConfigFile) 
		Set gblConfig = CreateObject("Microsoft.XMLDOM")
		gblConfig.Async = "False"
		gblConfig.Load( gblConfigFile )

		If gblConfig.xml = "" Then
			MsgBox "Failed to load " & gblConfigFile,, "XML Load Error"
		End If	
	End If
			' auto-load secondary configuration file (metadata), if any defined
'		Set xmlNode = gblConfig.selectSingleNode("/UFT/App/Metadata")
'		Dim metadataPath
'		If xmlNode is Nothing Then
'			metadataPath = ""	
'		Else
'			metadataPath = xmlNode.Text
'		End If
'
'		If Not StringIsNullOrEmpty(metadataPath) And IsFileExists(metadataPath) Then
'			Set gblConfig2 = CreateObject("Microsoft.XMLDOM")
'			gblConfig2.Async = "False"
'			gblConfig2.Load( metadataPath )
'			
'			If gblConfig2.xml = "" Then
'				MsgBox "Failed to load metadata " & metadataPath,, "XML Load Error"
'				Set GetConfigNode = Null
'				Exit Function
'			End If
'		End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Gets a collection of XML nodes from the global configuration file
'  strNodePath = Path to the node to be retrieved
' RETURNS: Collection of XmlNodes or NULL
' NOTE: Does not log any info to result report
Function GetConfigNodes(ByVal strNodePath)
	LoadConfigFile
	Set GetConfigNodes = gblConfig.selectNodes(strNodePath)			' search for node collection
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Gets a specified XML node from the global configuration file
'  strNodePath = Path to the node to be retrieved
' RETURNS: XmlNode or NULL
' NOTE: Does not log any info to result report
Function GetConfigNode(ByVal strNodePath)
	LoadConfigFile
	Set GetConfigNode = gblConfig.selectSingleNode(strNodePath)		' search for first node
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Gets the value for a specified node within the global configuration file
'  strNodePath = Path to the node 
' RETURNS: Value of the specified node
' NOTE: Does not log any info to result report
Function GetConfigValue(ByVal strNodePath)
	Dim arrOutput(), xmlNode, objNode, i

	Set xmlNode = GetConfigNode(strNodePath)
	If xmlNode is Nothing Then
		GetConfigValue = ""	
	Else
		GetConfigValue = xmlNode.Text
	End If
	Set xmlNode = Nothing
End Function


''' ************************************************
''' **** Secondary Helper Functions/Subroutines ****
''' ************************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Wait indefinitely until specified objects exists on screen.
'  objObject = Repository object to wait for
'  objNotUsed = Not used (FOR FUTURE USE)
' NOTE: Does not log any info to result report
Sub WaitUntilObjectExists(ByVal objObject, ByVal objNotUsed)
	Do
		If objObject.Exist(10) Then
			Exit Do
		End If
	Loop
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Gets a SQL connection string from the global configuration file. If not found, then searches Environment variables.
' RETURNS: SQL connection string for current server
' NOTE: Does not log any info to result report
Function GetConnectionString()
	Dim strServer : strServer = GetConfigValue("/UFT/DBServer/Name")
	Dim strConnect : strConnect = GetConfigValue("/UFT/DBServer/ConnectionString")
	
	If StringIsNullOrEmpty(strServer) Then
		On Error Resume Next
		strServer = Environment("DbServer_Name")
	End If
	If StringIsNullOrEmpty(strConnect) Then
		On Error Resume Next
		strConnect = Environment("DbServer_ConnectionString")
	End If
	
	GetConnectionString = Replace( strConnect, "$SERVER$", strServer)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Executes stored procedure defined in a specified XML node
'  objXmlNode = XML node containing stored procedure details
'  arrOutput = returned array to be loaded with output data 
' RETURNS: TRUE if successful; otherwise FALSE
' NOTE: Does not log any info to result report
Function ExecProcedure_GetArray(ByVal objXmlNode, ByRef arrOutput())
	Dim strSP : strSP = objXmlNode.getAttribute("name")
	Dim arrParams(), arrColumns(), colSubNodes, objSubNode, i

	ExecProcedure_GetArray = False

	' get parameters for procedure; if any
	Set colSubNodes = objXmlNode.selectNodes("InputParameter")
	i = 0

	For Each objSubNode in colSubNodes
		Dim strName : strName = objSubNode.getAttribute("name")
		Dim strValue : strValue = objSubNode.getAttribute("value")
		
		If Not IsEmpty(strName) and Not IsNull(strName) Then 'and Not IsEmpty(strValue) and Not IsNull(strValue) Then
			ReDim Preserve arrParams(1,i)
			arrParams(0,i) = strName
			arrParams(1,i) = strValue
		End If
	 	i = i + 1
	Next
	Set colSubNodes = Nothing
	
	' get columns to be retrieved; if any
	Set colSubNodes = objXmlNode.selectNodes("Column")
	ReDim Preserve arrColumns(-1)
	i = 0
	
	For Each objSubNode in colSubNodes
		ReDim Preserve arrColumns(i)
		arrColumns(i) = objSubNode.Text
	 	i = i + 1
	Next
	Set colSubNodes = Nothing

	' retrieve data from database
	ExecuteSQL GetConnectionString, strSP, arrParams, arrColumns, arrOutput
	ExecProcedure_GetArray = True
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get the number of rows in a data grid
'  objDataGrid = Datagrid respository object 
' RETURNS: Number of rows
' NOTE: Does not log any info to result report
Function GetDataGridRowCount(ByVal objDataGrid)
	GetDataGridRowCount = objDataGrid.RowCount
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get specified cell from a data grid (located at strColumnName and intRowIndex)
'  objDataGrid = Datagrid respository object 
'  strColumnName = Column name of the cell to retrieve
'  intRowIndex = 0-based index of the cell to retrieve
' RETURNS: Value of datagrid cell
' NOTE: Does not log any info to result report
Function GetDataGridCell(ByVal objDataGrid, ByVal strColumnName, ByVal intRowIndex)
	Dim intRowCount : intRowCount = objDataGrid.RowCount
	
	If (intRowIndex < 0) Or (intRowIndex > (intRowCount - 1)) Then
		GetDataGridCell = Empty
	End If

	Dim arrBuffer : arrBuffer = GetDataGridRow( objDataGrid, Array(strColumnName), intRowIndex)  

	GetDataGridCell = arrBuffer(0,0)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get specified row from a data grid (located at intRowIndex)
'  objDataGrid = Datagrid respository object 
'  arrColumns = Array of column names to retrieve
'  intRowIndex = 0-based index of the row to retrieve
' RETURNS: Array of data (column names requested X 1 row)
' NOTE: Does not log any info to result report
Function GetDataGridRow(ByVal objDataGrid, ByVal arrColumns, ByVal intRowIndex)
	Dim intRowCount : intRowCount = objDataGrid.RowCount
	Dim arrOutput(), i

	If (intRowIndex < 0) Or (intRowIndex > (intRowCount - 1)) Then
		GetDataGridRow = arrOutput
	End If

	Dim arrBuffer : arrBuffer = GetDataGridData( objDataGrid, arrColumns)
	ReDim Preserve arrOutput(Ubound(arrColumns), 0)

	For i = 0 To Ubound(arrColumns)	' each requested columnName
		arrOutput(i,0) = arrBuffer(i, intRowIndex)	' copy rows for matching column
	Next

	GetDataGridRow = arrOutput
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get specified rows from a data grid 
'  objDataGrid = Datagrid respository object 
'  arrColumns = Array of column names to retrieve
'  arrRowIndexes = Array of 0-based row indexes to retrieve
' RETURNS: Array of data (column names requested X N rows)
' NOTE: Does not log any info to result report
Function GetDataGridRows(ByVal objDataGrid, ByVal arrColumns, ByVal arrRowIndexes)
	Dim intRowCount : intRowCount = objDataGrid.RowCount
	Dim arrOutput(), i,j

	Dim arrBuffer : arrBuffer = GetDataGridData( objDataGrid, arrColumns)
	ReDim Preserve arrOutput(Ubound(arrColumns), UBound(arrRowIndexes))

	For J = 0 To Ubound(arrRowIndexes)
		For i = 0 To Ubound(arrColumns)	' each requested columnName
			arrOutput(i,j) = arrBuffer(i, arrRowIndexes(j))	' copy rows for matching column
		Next		
	Next

	GetDataGridRows = arrOutput
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get specified columns from a data grid
'  objDataGrid = Datagrid respository object 
'  strDataGridName = Name of the data grid repository object
'  arrColumns = Array of column names to retrieve
' RETURNS: Array of data (column names requested X all rows in grid)
' NOTE: Does not log any info to result report
Function GetDataGridData(ByVal objDataGrid, ByVal arrColumns)
	Dim intRowCount : intRowCount = objDataGrid.RowCount
	Dim intColCount : intColCount = objDataGrid.ColumnCount
	Dim arrOutput(), intRequest, i, j

	ReDim Preserve arrOutput(Ubound(arrColumns), intRowCount-1)

	For intRequest = 0 To Ubound(arrColumns)	' each requested columnName
		For i = 0 To intColCount-1				' scan table columns
			If objDataGrid.GetCellProperty(0, i, "colname") = arrColumns(intRequest) Then	' found matching column
				For j = 0 To intRowCount-1
					arrOutput(intRequest,j) = objDataGrid.GetCellData(j ,i)	' copy rows for matching column
				Next
			End If
		Next
	Next

	GetDataGridData = arrOutput
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get entire contents from a data grid 
'  objDataGrid = Datagrid respository object 
'  isGetAllColumns = TRUE to retrieve all columns (including hidden columns); otherwise FALSE to get only visible columns
'  arrColumnNames = Returns an array of the column names retrieved
' RETURNS: Array of data
' NOTE: Does not log any info to result report
Function GetDataGridContents(ByVal objDataGrid, ByVal isGetAllColumns, ByRef arrColumnNames())
	Dim intRowCount : intRowCount = objDataGrid.RowCount
	Dim intColCount : intColCount = objDataGrid.ColumnCount
	Dim intVisibleColumns : intVisibleColumns = 0
	Dim arrOutput(), i, j, k

	' calc number of columns to return
	If isGetAllColumns Then
		intVisibleColumns = intColCount
	Else
		For i = 0 To intColCount-1
			If objDataGrid.Object.Columns.Item(i).Visible Then
				intVisibleColumns = intVisibleColumns + 1
			End If
		Next
	End If

	' resize output arrays
	ReDim arrOutput(intVisibleColumns-1, intRowCount-1)
	ReDim arrColumnNames(intVisibleColumns-1)

	' get data
	If isGetAllColumns Then	' get all columns (including those hidden from view)
		For i = 0 To intColCount-1											' all columns
			arrColumnNames(i) = objDataGrid.GetCellProperty(0, i, "colname")
			For j = 0 To intRowCount-1										' all rows
				arrOutput(i,j) = objDataGrid.GetCellData(j ,i)
			Next
		Next
	Else					' get visible columns	
		k = 0
		For i = 0 To intColCount-1											' scan all columns
			If objDataGrid.Object.Columns.Item(i).Visible Then					' only visible columns
				arrColumnNames(k) = objDataGrid.GetCellProperty(0, i, "colname")
				For j = 0 To intRowCount-1									' all rows
					arrOutput(k,j) = objDataGrid.GetCellData(j ,i)
				Next
				k = k + 1
			End If
		Next
	End If
	
	GetDataGridContents = arrOutput
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get the maximum value for a specified column in a data grid
'  objDataGrid = Datagrid respository object 
'  strColumnName = Name of the column (e.g. column header)
' RETURNS: Maximum string value found
' NOTE: Does not log any info to result report
Function GetDataGridColumnMax(ByVal objDataGrid, ByVal strColumnName)
	Dim arrBuffer : arrBuffer = GetDataGridData(objDataGrid, Array(strColumnName))
	Dim strMax : strMax = ""
	Dim strValue
	
	For Each strValue in arrBuffer
		If strValue > strMax Then
			strMax = strValue
		End If
	Next

	GetDataGridColumnMax = strMax
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get the minimum value for a specified column in a data grid
'  objDataGrid = Datagrid respository object 
'  strColumnName = Name of the column (e.g. column header)
' RETURNS: Minimum string value found
' NOTE: Does not log any info to result report
Function GetDataGridColumnMin(ByVal objDataGrid, ByVal strColumnName)
	Dim arrBuffer : arrBuffer = GetDataGridData(objDataGrid, Array(strColumnName))
	Dim strMin : strMin = ""
	Dim strValue
	
	For Each strValue in arrBuffer
		If strValue < strMin Then
			strMin = strValue
		End If
	Next

	GetDataGridColumnMin = strMin
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Selects the first row in a data grid that contains a specified value in a specified column
'  objDataGrid = Datagrid respository object 
'  strColumnName = Name of the column to search
'  strSearchValue = Value of the text to search for
'  isRegEx = TRUE if search value is regular expression; otherwise FALSE for normal text search
' RETURNS: Zero-based row index; -1 if not found
' NOTE: Does not log any info to result report
Function FindRowInDataGrid( ByVal objDataGrid, ByVal strColumnName, ByVal strSearchValue, ByVal isRegEx)
	' gets all values in the column
	Dim arrData : arrData = GetDataGridData(objDataGrid, Array(strColumnName) )
	' locates the search text
	FindRowInDataGrid = FindInArray(arrData, strSearchValue, isRegEx, null, null)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Selects the first row in a data grid that contains a specified value in a specified column
'  objDataGrid = Datagrid respository object 
'  strColumnName = Name of the column to search
'  strSearchValue = Value of the text to search for
'  isRegEx = TRUE if search value is regular expression; otherwise FALSE for normal text search
' NOTE: Does not log any info to result report
Sub SelectRowInDataGrid( ByVal objDataGrid, ByVal strColumnName, ByVal strSearchValue, ByVal isRegEx)
	' locates the row
	Dim intRowIndex : intRowIndex = FindRowInDataGrid( objDataGrid, strColumnName, strSearchValue, isRegEx)
	' highlights the row, if found
	If intRowIndex > -1 Then
		objDataGrid.SelectRow intRowIndex		
	End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Sets the checkbox of a specified node within a treeview
'  objTree = Treeview repository object
'  arrNodeNames = List containing the path to the specified node (e.g. array with node path hierarchy)
'  isChecked = TRUE to set last node in the specified path to Checked/Checkmark; otherwise FALSE to clear last node in the specified path
' NOTE: Will log info to result report
Sub TreeViewSetCheckbox(ByVal objTree, ByVal arrNodeNames, ByVal isChecked)
	Dim objNode, strName, strPath, intNodeCount, i
	Set objNode = objTree.Object

	For Each strName in arrNodeNames
		If Len(strPath) = 0 Then
			strPath = strName
		Else			
			strPath = strPath & ";" & strName
		End If
		intNodeCount = objNode.Nodes.Count
		
		For  i = 0 to intNodeCount - 1
			If strName = objNode.Nodes.Item(i).Text Then
		    	Set objNode = objNode.Nodes.Item(i)
		    	Exit For
			End If
		Next
	
		If i = intNodeCount Then
			Set objNode = Nothing
			Exit For
		End If
	Next

	If Not objNode Is Nothing Then
		objNode.Checked = isChecked
		If isChecked Then
			reporter.ReportEvent micPass, "TreeViewSetCheckbox", "Set checkbox at path (" & strPath & ") to Checked state"
		Else
			reporter.ReportEvent micPass, "TreeViewSetCheckbox", "Set checkbox at path (" & strPath & ") to Unchecked state"
		End If
	Else
		reporter.ReportEvent micFail, "TreeViewSetCheckbox", "Invalid path (" & strPath & ")"
	End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Find a specified node within a treeview
'  objTree = Treeview repository object
'  arrNodeNames = List containing the path to the specified node (e.g. array with node path hierarchy)
' RETURN: TRUE if node is found; otherwise FALSE if node is not found
' NOTE: Does not log any info to result report
Function TreeViewFindNode(ByVal objTree, ByVal arrNodeNames)
	Dim objNode, strName, strPath, intNodeCount, i
	Set objNode = objTree.Object

	For Each strName in arrNodeNames
		If Len(strPath) = 0 Then
			strPath = strName
		Else			
			strPath = strPath & ";" & strName
		End If
		intNodeCount = objNode.Nodes.Count
		
		For  i = 0 to intNodeCount - 1
			If strName = objNode.Nodes.Item(i).Text Then
		    	Set objNode = objNode.Nodes.Item(i)
		    	Exit For
			End If
		Next
	
		If i = intNodeCount Then
			Set objNode = Nothing
			Exit For
		End If
	Next

	If objNode Is Nothing Then
		TreeViewFindNode = False
	Else
		TreeViewFindNode = True
	End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Sets a value in a multi-line text box
'  objField = Multi-line text box object
'  strValue = value to set in text box
' NOTE: Does not log any info to result report
Sub SetMultiTextBox(ByVal objField, ByVal strValue)
	objField.SetCaretPos 0,0
	While Len(objField.GetROProperty("Text")) <> 0	'DELETE any existing value
		objField.Type micDel
	Wend

	objField.Type strValue	' set new value
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Sets a value in a multi-line text box
'  objWindow = window repository object
'  strName = name of field object
'  strType = type of control
'  strValue = value to set in text box
' NOTE: Does not log any info to result report
Sub SetFieldOnScreen(ByVal objWindow, ByVal strName, ByVal strType, ByVal strValue)
	Select Case LCase(strType)
		Case "textbox"
			objWindow.SwfEdit(strName).Set strValue
		Case "dropdown"
			objWindow.SwfComboBox(strName).Select strValue
		Case "numberspinner"
			objWindow.SwfSpin(strName).Set CInt(strValue)
		Case "checkbox"
			If (LCase(strValue) = "on") Or (LCase(strValue) = "true") Or (strValue = "1") Then
				objWindow.SwfCheckBox(strName).Set "ON"
			Else		
				objWindow.SwfCheckBox(strName).Set "OFF"
			End If
		Case "radiobutton"
			If (LCase(strValue) = "on") Or (LCase(strValue) = "true") Or (strValue = "1") Then
				objWindow.SwfRadioButton(strName).Set
			End If
		Case "multilinetextbox"
			SetMultiTextBox objWindow.SwfEditor(strName), strValue
		Case "calendar"
			objWindow.SwfCalendar(strName).SetDate strValue
		Case "calendartime"
			objWindow.SwfCalendar(strName).SetTime strValue
		Case "calendarrange"
			objWindow.SwfCalendar(strName).SetDateRange strValue
		Case "winedit"
			objWindow.WinEdit(strName).Set strValue
		Case Else
			print Now() & " - ERROR - UNKNOWN FIELD TYPE IGNORED " & strType
	End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' END OF SCRIPT
