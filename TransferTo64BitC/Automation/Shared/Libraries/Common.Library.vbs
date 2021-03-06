
' HEADER
'------------------------------------------------------------------'
'    Description:  General Common Library                          '
'                  (e.g. SQL, Arrays, RegEx, Files, Folders, other)'
'                                                                  '
'        Project:  UFT Automation                                  '
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
'  20170511  v2.4      RNiedzwiecki  Added functions FileCount, MoveFolder
'                                    Changed subroutine DeleteFolder to force delete
'                                    Change ExecuteSQL to handle sql that creates/drops a procedure
'  20170428  v2.3      RNiedzwiecki  Replaced subroutine AddToResults with expanded LogResult
'  20170426  v2.2      RNiedzwiecki  Added hash-separated-value (HSV) output to AddToResults
'  20170421  v2.1.1    RNiedzwiecki  Corrected CString default value
'  20170301  v2.1      RNiedzwiecki  Added subroutine AppendFile, AddToResults
'  20170201  v2.0      RNiedzwiecki  CODE SPLIT from Guardian Library.qfl; this file only contains common and generic functionality
'                                    Added subroutine PurgeDateTimeSeconds 
'                                    Added function ProjectWorkingPath
'                                    Changed ExecuteSQL to no longer strip seconds (use PurgeDateTimeSeconds instead)
'  20170131  v1.2      RNiedzwiecki  Added subroutine SetConfigFile
'                                    Expanded GetConfigNode to support Environment variables (File>Settings>Environment>UserDefined)
'                                    Added support for secondary configuration with metadata
'                                    Added timestamps to Print debug statements
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
'  NA

' START SCRIPT
Option Explicit

''' **************************************
''' **** String Functions/Subroutines ****
''' **************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Indicates whether a specified string is null or empty or blank
'  strValue = String to be evaluated
' RETURN: TRUE if value has no text; otherwise FALSE if value has some text
' NOTE: Does not log any info to result report
Function StringIsNullOrEmpty(ByVal strValue)
   	If IsNull(strValue) Then
   		StringIsNullOrEmpty = True
   		Exit Function
   	End If
   	If IsEmpty(strValue) Then
   		StringIsNullOrEmpty = True
   		Exit Function
   	End If
   	If strValue = "" Then
   		StringIsNullOrEmpty = True
   		Exit Function
   	End If

	StringIsNullOrEmpty = False
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Indicates whether the beginning of a specified string matches another string
'  strStringToBeSearched = String to be evaluated
'  strStringToSearchFor = String to search for
' RETURN: TRUE if the beginning of the strings match; otherwise FALSE if they do not match
' NOTE: Does not log any info to result report
Function StringStartsWith(ByVal strStringToBeSearched, ByVal strStringToSearchFor )
	StringStartsWith = False
	If Left(strStringToBeSearched, Len(strStringToSearchFor)) = strStringToSearchFor Then
		StringStartsWith = True
	End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Indicates whether the ending of a specified string matches another string
'  strStringToBeSearched = String to be evaluated
'  strStringToSearchFor = String to search for
' RETURN: TRUE if the ending of the strings match; otherwise FALSE if they do not match
' NOTE: Does not log any info to result report
Function StringEndsWith(ByVal strStringToBeSearched, ByVal strStringToSearchFor )
	StringEndsWith = False
	If Right(strStringToBeSearched, Len(strStringToSearchFor)) = strStringToSearchFor Then
		StringEndsWith = True
	End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Converts a string value to a string (equivalent to CStr, but handles null values also)
'  strValue = String value to be converted to string
' RETURN: String equivalent of the specified value
' NOTE: Does not log any info to result report
Function CString(ByVal strValue)
	CString = null
	On Error Resume Next
	CString = CStr(strValue)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Converts a string value to a boolean (equivalent to CBool, but handles null values also)
'  strValue = String value to be converted to boolean
' RETURN: Boolean equivalent of the specified value
' NOTE: Does not log any info to result report
Function CBoolean(ByVal strValue)
	CBoolean = false
	On Error Resume Next
	CBoolean = CBool(strValue)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Converts a string value to an integer (equivalent to CInt, but handles null values also)
'  strValue = String value to be converted to a number
' RETURN: Integer equivalent of the specified value
' NOTE: Does not log any info to result report
Function CInteger(ByVal strValue)
	CInteger = null
	On Error Resume Next
	CInteger = CInt(strValue)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Converts a string value to a long (equivalent to CLng, but handles null values also)
'  strValue = String value to be converted to a number
' RETURN: Long equivalent of the specified value
' NOTE: Does not log any info to result report
Function CLong(ByVal strValue)
	CLong = null
	On Error Resume Next
	CLong = CLng(strValue)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Converts a string of delimited values into an array
'  strValue = String to be parsed 
'  objDelimiter = Delimiter value used for parsing
' RETURN: Array of parsed values
' NOTE: Does not log any info to result report
Function CArray(ByVal strValue, ByVal objDelimiter)
	CArray = null
	On Error Resume Next	
	CArray = Split(strValue, objDelimiter)
End Function


''' ***********************************
''' **** SQL Functions/Subroutines ****
''' ***********************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Executes SQL in a database
'  strConnect = SQL connection string
'  strSQL = SQL statement or SQL stored procedure name [procedure name must contain "usp_"]
'  arrInputParams = array of input parameters for stored procedure [null=no parameters]
'                   recommend 2D list of parameter name/value pairs 
'                   alternate option is 1D list starting with null followed by sequential list of parameter values
'  arrOutputFields = list of column names to be retrieved from recordset [null=all columns]
'  arrOutput = returned array to be loaded with output data 
' RETURNS: TRUE if successful; otherwise FALSE
' NOTE: Does not log any info to result report
Function ExecuteSQL(ByVal strConnect, ByVal strSQL, ByVal arrInputParams, ByVal arrOutputFields, ByRef arrOutput())
	Dim conn, cmd, rs

	' set the connection
	Set conn = CreateObject("ADODB.Connection")
	conn.ConnectionString = strConnect
	conn.Open 
	
	' set the command
	SET cmd = CreateObject("ADODB.Command")
	SET cmd.ActiveConnection = conn
	cmd.CommandText = strSQL
	If InStr(LCase(Left(strSQL, 50)), "usp_") > 0 Then
		cmd.CommandType = 4 'adCmdStoredProc
	Else
		cmd.CommandType = 1 'adCmdText
	End If

	' set the optional parameters
	If IsArray(arrInputParams) And (NumberOfDimensions(arrInputParams) > -1) Then ' any input parameters
		Dim i
		If NumberOfDimensions(arrInputParams) < 2 Then		' 1D list
			cmd.Parameters.Refresh
			
			For i = 0 To Ubound(arrInputParams)	
				If Not IsNull(arrInputParams(i)) Then		' parameter(0) should be null and ignored
					cmd.Parameters(i) = arrInputParams(i)	' set by parameter index
				End If
			Next	
		End If
		If NumberOfDimensions(arrInputParams) = 2 Then		' 2D list of name/value pairs	
			cmd.Parameters.Refresh
			
			For i = 0 To Ubound(arrInputParams,2)		
				cmd.Parameters(arrInputParams(0,i)) = arrInputParams(1,i)	' set by parameter name
			Next	
		End If
	End If
	
	' execute the command
	SET rs = cmd.Execute

	If IsNull(arrOutput) Then
		conn.Close
		SET cmd = Nothing
		SET conn = Nothing
		ExecuteSQL = False
		Exit Function
	End If

	' copy the data to the output array
	If rs.EOF Then 
		ExecuteSQL = False
	Else 
		Dim col, row, arrParts, intCount
		row = 0
		
		If NumberOfDimensions(arrOutputFields) = 0 Then		' get all columns, if undimensioned array
			intCount = rs.Fields.Count
			ReDim arrOutput(intCount-1, -1)
			
			Do While NOT rs.Eof   
				ReDim Preserve arrOutput(intCount-1, row)
	
				For col = 0 To (intCount-1)
					arrOutput(col,row) = rs.Fields.Item(col)
					
'					If IsDate(arrOutput(col, row)) And (InStr(arrOutput(col,row), ":") > 0) Then		' strip out the seconds from date/time columns
'						arrParts  = Split(arrOutput(col, row), ":")
'						arrOutput(col, row) = arrParts(0) & ":" & arrParts(1) & " " &  Right(arrParts(2),2)
'					End If
				Next	
	
				rs.MoveNext     
				row = row + 1
			Loop
		Else
			If Ubound(arrOutputFields) > 0 Then				' get specified columns
				ReDim arrOutput(Ubound(arrOutputFields), -1)
				
				Do While NOT rs.Eof   
					ReDim Preserve arrOutput(Ubound(arrOutputFields), row)
		
					For col = 0 To Ubound(arrOutputFields)
						arrOutput(col,row) = rs(arrOutputFields(col))

'						If IsDate(arrOutput(col, row)) And (InStr(arrOutput(col,row), ":") > 0) Then	' strip out the seconds from date/time columns
'							arrParts  = Split(arrOutput(col, row), ":")
'							arrOutput(col, row) = arrParts(0) & ":" & arrParts(1) & " " &  Right(arrParts(2),2)
'						End If
					Next	
		
					rs.MoveNext     
					row = row + 1
				Loop
			ElseIf Ubound(arrOutputFields) = 0 Then			' get single column
				ReDim arrOutput(-1)

				Do While NOT rs.Eof   
					ReDim Preserve arrOutput(row)
					arrOutput(row) = rs(arrOutputFields(0))
		
'						If IsDate(arrOutput(row)) And (InStr(arrOutput(row), ":") > 0)  Then		' strip out the seconds from date/time columns
'							arrParts = Split(arrOutput(row), ":")
'							arrOutput(row) = arrParts(0) & ":" & arrParts(1) & " " &  Right(arrParts(2),2)
'						End If
					rs.MoveNext     
					row = row + 1
				Loop
			ElseIf Ubound(arrOutputFields) = -1 Then		' get all columns, if empty array
				intCount = rs.Fields.Count
				ReDim arrOutput(intCount-1, -1)
				
				Do While NOT rs.Eof   
					ReDim Preserve arrOutput(intCount-1, row)
		
					For col = 0 To (intCount-1)
						arrOutput(col,row) = rs.Fields.Item(col)
						
'						If IsDate(arrOutput(col, row)) And (InStr(arrOutput(col,row), ":") > 0) Then		' strip out the seconds from date/time columns
'							arrParts  = Split(arrOutput(col, row), ":")
'							arrOutput(col, row) = arrParts(0) & ":" & arrParts(1) & " " &  Right(arrParts(2),2)
'						End If
					Next	
		
					rs.MoveNext     
					row = row + 1
				Loop
			End If		
		End If
		
		ExecuteSQL = True
	End If

	' cleanup
	rs.Close
	conn.Close
	SET rs = Nothing
	SET cmd = Nothing
	SET conn = Nothing
End Function


''' *************************************
''' **** Array Functions/Subroutines ****
''' *************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Searches an array for the specified value
'  arrArray = Array to search
'  objValue = Value to search for
'  isValueRegEx = TRUE if search value is RegEx mask; otherwise FALSE if search value is non-RegEx
'  intColumnIndex = Zero-based index of column to search [NULL=search all columns]
'  intRowIndex = Zero-based index of row to search [NULL=search all rows]
' RETURNS: TRUE if value found in the array; otherwise FALSE if value was not found
' NOTE: Will log info to result report
Function IsValueInArray( ByVal arrArray, ByVal objValue, ByVal isValueRegEx, ByVal intColumnIndex, ByVal intRowIndex )
	Dim index : index = FindInArray( arrArray, objValue, isValueRegEx, intColumnIndex, intRowIndex )
	If index = -1 Then
		IsValueInArray = False
	Else
		IsValueInArray = True
	End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Locates a value in the specified array
'  arrArray = Array to search
'  objValue = Value to search for
'  isValueRegEx = TRUE if search value is RegEx mask; otherwise FALSE for regular value match
'  intColumnIndex = Zero-based index of column to search [NULL=search all columns]
'  intRowIndex = Zero-based index of row to search [NULL=search all rows]
' RETURNS: Zero-based index where value was found [-1 = not found]
' NOTE: Will log info to result report
Function FindInArray( ByVal arrArray, ByVal objValue, ByVal isValueRegEx, ByVal intColumnIndex, ByVal intRowIndex )
	Dim strStep : strStep = "FindInArray"
	Dim i, j
	FindInArray = -1	' not found

	' verify parameters are valid
	If IsNull(arrArray) Or IsEmpty(arrArray) Then
		reporter.ReportEvent micFail, strStep, "Invalid parameter: array"
		Exit Function
	End If
	If Not IsArray(arrArray) Then
		reporter.ReportEvent micFail, strStep, "Parameter is not an Array"
		Exit Function
	End If
	If IsNull(objValue) Or IsEmpty(objValue) Then
		reporter.ReportEvent micFail, strStep, "Invalid parameter: value"
		Exit Function
	End If

	If NumberOfDimensions(arrArray) = 1 Then 
		' 1D arrays (a.k.a. single dimensional list)   ignore column/row indexes and just search the list provided
		For i = 0 To Ubound(arrArray, 1)
			If isValueRegEx Then
				If RegExpTest(objValue, arrArray(i), false) Then
					reporter.ReportEvent micDone, strStep, "Found value at index " & i				
					FindInArray = i
					Exit Function
				End If
			Else									
				If arrArray(i) = objValue Then			' exact match
					reporter.ReportEvent micDone, strStep, "Found value at index " & i				
					FindInArray = i
					Exit Function
				End If
			End If
		Next
	Else	
		' 2D arrays     search specified indexes, if any
		If (IsNull(intRowIndex) Or IsEmpty(intRowIndex)) And (IsNull(intColumnIndex) Or IsEmpty(intColumnIndex)) Then	' anywhere in array
			Dim count : count = -1
			For i = 0 To Ubound(arrArray, 1)
				For j = 0 To Ubound(arrArray, 2)
					count = count + 1
					If isValueRegEx Then
						If RegExpTest(objValue, arrArray(i, j), false) Then
							reporter.ReportEvent micDone, strStep, "Found value at index (" & i & "," & j & ")"
							FindInArray = count
							Exit Function
						End If
					Else
						If arrArray(i, j) = objValue Then				' exact match
							reporter.ReportEvent micDone, strStep, "Found value at index (" & i & "," & j & ")"
							FindInArray = count
							Exit Function
						End If
					End If
				Next
			Next	
		ElseIf IsNull(intRowIndex) Or IsEmpty(intRowIndex) Then			' for specified column
			For i = 0 To Ubound(arrArray, 2)							' search rows
				If isValueRegEx Then					
					If RegExpTest(objValue, arrArray(intColumnIndex, i), false) Then
						reporter.ReportEvent micDone, strStep, "Found value at index (" & intColumnIndex & "," & i & ")"
						FindInArray = i
						Exit Function
					End If
				Else
					If arrArray(intColumnIndex, i) = objValue Then		' exact match
						reporter.ReportEvent micDone, strStep, "Found value at index (" & intColumnIndex & "," & i & ")"
						FindInArray = i
						Exit Function
					End If
				End If
			Next	
		ElseIf IsNull(intColumnIndex) Or IsEmpty(intColumnIndex) Then	' for specified row
			For i = 0 To Ubound(arrArray, 1)							' search column
				If isValueRegEx Then
					If RegExpTest(objValue, arrArray(i, intRowIndex), false) Then
						reporter.ReportEvent micDone, strStep, "Found value at index (" & i & "," & intRowIndex & ")"
						FindInArray = i
						Exit Function
					End If
				Else					
					If arrArray(i, intRowIndex) = objValue Then			' exact match
						reporter.ReportEvent micDone, strStep, "Found value at index (" & i & "," & intRowIndex & ")"
						FindInArray = i
						Exit Function
					End If
				End If
			Next
		End If
	End If
	
	reporter.ReportEvent micDone, strStep, "Value not found in array"
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Checks if two arrays match each other
'  arrA = First array 
'  arrB = Second array
'  isValueComparison = TRUE to perform value to value comparison using original value type; otherwise FALSE to compare data as string text
'  strComment = Optional comment to include in report
' NOTE: Will log info to result report
Sub CheckIfArraysMatch( ByVal arrA, ByVal arrB, ByVal isValueComparison, ByVal strComment)
	Dim strStep : strStep = "CheckIfArraysMatch" & "-" & strComment
	Dim i, j

	' verify parameters are valid arrays
	If Not IsArray(arrA) Then
		reporter.ReportEvent micFail, strStep, "First parameter is not an Array"
		Exit sub	
	End If
	If Not IsArray(arrB) Then
		reporter.ReportEvent micFail, strStep, "Second parameter is not an Array"
		Exit sub	
	End If

	' verify arrays have same dimensions
	If NumberOfDimensions(arrA) <> NumberOfDimensions(arrB) Then
		reporter.ReportEvent micFail, strStep, "Mismatched dimensions; Array1 is " & NumberOfDimensions(arrA) & "D and Array2 is " & NumberOfDimensions(arrB) & "D"
		Exit sub	
	End If

	' verify first dimension has matching quantity
	If Ubound(arrA, 1) <> Ubound(arrB, 1) Then
		reporter.ReportEvent micFail, strStep, "Mismatched quantity in dimension #1; Array1 has " & Ubound(arrA, 1)+1 & " items and Array2 has " & Ubound(arrB, 1)+1
		Exit sub	
	End If
	
	' verify second dimension has matching quanity; if second dimension exists
	If (NumberOfDimensions(arrA) = 2) And (NumberOfDimensions(arrA) = 2) Then
		If Ubound(arrA, 2) <> Ubound(arrB, 2) Then
			reporter.ReportEvent micFail, strStep, "Mismatched quantity in dimension #2; Array1 has " & Ubound(arrA, 2)+1 & " items and Array2 has " & Ubound(arrB, 2)+1
			Exit sub	
		End If		
	End If	

	' compare arrays
	If (NumberOfDimensions(arrA) = 1) And (NumberOfDimensions(arrB) = 1) Then 
		' 1D arrays (a.k.a. single dimensional lists)
		For i = 0 To Ubound(arrA, 1)
			If isValueComparison Then			
				If arrA(i) <> arrB(i) Then
					reporter.ReportEvent micFail, strStep, "Value mismatch at item(" & i & ") value=" & arrA(i) & " of type=" & VarType(arrA(i)) & " vs value=" & arrB(i) & " of type=" & VarType(arrB(i))
					Exit sub
				End If
			Else	' compare as text 
				If CStr(arrA(i)) <> CStr(arrB(i)) Then
					reporter.ReportEvent micFail, strStep, "Visual mismatch at item(" & i & ") value=" & arrA(i) & " vs value=" & arrB(i) 
					Exit sub
				End If
			End If
		Next
	Else	
		' 2D arrays
		For i = 0 To Ubound(arrA, 1)
			For j = 0 To Ubound(arrA, 2)
				If isValueComparison Then
					If arrA(i,j) <> arrB(i,j) Then
						reporter.ReportEvent micFail, strStep, "Value mismatch at coordinate(" & i & "," & j & ") value=" & arrA(i,j) & " of type=" & VarType(arrA(i,j)) & " vs value=" & arrB(i,j) & " of type=" & VarType(arrB(i,j))
						Exit sub
					End If
				Else	' compare as text 
					If CStr(arrA(i,j)) <> CStr(arrB(i,j)) Then
						reporter.ReportEvent micFail, strStep, "Visual mismatch at coordinate(" & i & "," & j & ") value=" & arrA(i,j) & " vs value=" & arrB(i,j)
						Exit sub
					End If
				End If
			Next
		Next
	End If

	reporter.ReportEvent micPass, strStep, "Arrays match"
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Purge rows from specified array
'  arrArray = Array to purge
'  intIndex = Zero-based index of first row to purge
'  intLength = Number of rows to purge
Function PurgeArrayRows( ByVal arrArray, ByVal intIndex, ByVal intLength)
	PurgeArrayRows = Null

	' verify array is valid
	If Not IsArray(arrArray) Then
		reporter.ReportEvent micFail, "PurgeArrayRows", "First parameter is not an Array"
		Exit Function	
	End If

	' ignore purge process if nothing to purge
	If intLength = 0 Then
		reporter.ReportEvent micPass, "PurgeArrayRows", "Length = 0; Nothing to purge, original array returned"
		PurgeArrayRows = arrArray
		Exit Function
	End If

	Dim i, j, row, col, arrOutput

	' purge arrays
	If NumberOfDimensions(arrArray) = 1 Then  
		' 1D arrays (a.k.a. single dimensional lists)

		' verify valid index
		If (intIndex < 0) Or (intIndex > Ubound(arrArray)) Then
			reporter.ReportEvent micFail, "PurgeArrayRows", "Invalid index[" & intIndex & "]; Array has " & Ubound(arrArray)+1 & " items"
			Exit Function	
		End If
	
		' verify valid length
		If (intLength < 0) Or (intLength > (Ubound(arrArray)+1)) Then
			reporter.ReportEvent micFail, "PurgeArrayRows", "Invalid length[" & intLength & "]; Array has " & Ubound(arrArray)+1 & " items"
			Exit Function	
		End If

		ReDim arrOutput(UBound(arrArray)-intLength)	' shrunken output array
		row = 0

		' copy rows before the index
		For i = 0 To (intIndex - 1)
			arrOutput(row) = arrArray(i)
			row = row + 1
		Next
		' copy rows after the purged gap
		For i = (intIndex + intLength) To Ubound(arrArray)
			arrOutput(row) = arrArray(i)
			row = row + 1
		Next
	Else	
		' 2D arrays
		
		' verify valid index
		If (intIndex < 0) Or (intIndex > Ubound(arrArray, 2)) Then
			reporter.ReportEvent micFail, "PurgeArrayRows", "Invalid index[" & intIndex & "]; Array has " & Ubound(arrArray, 2)+1 & " items"
			Exit Function	
		End If
	
		' verify valid length
		If (intLength < 0) Or (intLength > (Ubound(arrArray, 2)+1)) Then
			reporter.ReportEvent micFail, "PurgeArrayRows", "Invalid length[" & intLength & "]; Array has " & Ubound(arrArray, 2)+1 & " items"
			Exit Function	
		End If
		
		ReDim arrOutput(UBound(arrArray,1), UBound(arrArray,2)-intLength)	' shrunken row output
		col = 0

		For i = 0 To Ubound(arrArray, 1)	' each column
			row = 0
			
			' copy rows before the index
			For j = 0 To (intIndex - 1)		
				arrOutput(col,row) = arrArray(i,j)
				row = row + 1
			Next
			' copy rows after the purged gap
			For j = (intIndex + intLength) To Ubound(arrArray, 2)
				arrOutput(col,row) = arrArray(i,j)
				row = row + 1
			Next
			
			col = col + 1	' next column
		Next
	End If

	PurgeArrayRows = arrOutput
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Purge columns from specified array
'  arrArray = Array to purge
'  intIndex = Zero-based index of first column to purge
'  intLength = Number of columns to purge
Function PurgeArrayColumns( ByVal arrArray, ByVal intIndex, ByVal intLength)
	PurgeArrayColumns = Null
		
	' verify array is valid
	If Not IsArray(arrArray) Then
		reporter.ReportEvent micFail, "PurgeArrayColumns", "First parameter is not an Array"
		Exit Function	
	End If

	' ignore purge process if nothing to purge
	If intLength = 0 Then
		reporter.ReportEvent micPass, "PurgeArrayColumns", "Length = 0; Nothing to purge, original array returned"
		PurgeArrayColumns = arrArray
		Exit Function
	End If

	Dim i, j, row, col, arrOutput

	' purge arrays
	If NumberOfDimensions(arrArray) = 1 Then  
		' 1D arrays (a.k.a. single dimensional lists)
		PurgeArrayColumns = PurgeArrayRows(arrArray, intIndex, intLength) 	' implemented in other function
		Exit Function
	Else	
		' 2D arrays
		
		' verify valid index
		If (intIndex < 0) Or (intIndex > Ubound(arrArray, 1)) Then
			reporter.ReportEvent micFail, "PurgeArrayColumns", "Invalid index[" & intIndex & "]; Array has " & Ubound(arrArray, 1)+1 & " items"
			Exit Function	
		End If
	
		' verify valid length
		If (intLength < 0) Or (intLength > (Ubound(arrArray, 1)+1)) Then
			reporter.ReportEvent micFail, "PurgeArrayColumns", "Invalid length[" & intLength & "]; Array has " & Ubound(arrArray, 1)+1 & " items"
			Exit Function	
		End If
		
		ReDim arrOutput(UBound(arrArray,1)-intLength, UBound(arrArray,2))	' shrunken column output
		row = 0

		For j = 0 To Ubound(arrArray, 2)	' each row
			col = 0
			
			' copy columns before the index
			For i = 0 To (intIndex - 1)		
				arrOutput(col,row) = arrArray(i,j)
				col = col + 1
			Next
			' copy columns after the purged gap
			For i = (intIndex + intLength) To Ubound(arrArray, 1)
				arrOutput(col,row) = arrArray(i,j)
				col = col + 1
			Next
			
			row = row + 1	' next row
		Next
	End If

	PurgeArrayColumns = arrOutput
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Gets the number of dimensions in the specified array
'  arrArray = Array to be scanned
' RETURNS: Number of dimensions in the specified array
' NOTE: Does not log any info to result report
Function NumberOfDimensions(ByVal arrArray)
    Dim intDimensions
    On Error Resume Next
	
	intDimensions = 0
    
    Do While Err.number = 0
        intDimensions = intDimensions + 1
        UBound arrArray, intDimensions
    Loop
    
    On Error Goto 0
    NumberOfDimensions = intDimensions - 1
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Strips date, time or seconds in the specified array
'  arrArray = Array to be purged 
'  isPurgeDate = TRUE if date shall be purged; otherwise FALSE to keep date
'  isPurgeTime = TRUE if time shall be purged; otherwise FALSE to keep time
'  isPurgeSeconds = TRUE if seconds shall be purged; otherwise FALSE to keep seconds
'  isString = TRUE to convert purged cells to string type; otherwise FALSE to keep purged cells as date type
' NOTE: Will log info to result report
Sub PurgeDateTimeSeconds( ByRef arrArray, ByVal isPurgeDate, ByVal isPurgeTime, ByVal isPurgeSeconds, ByVal isString)
	Dim strStep : strStep = "PurgeDateTime"
	Dim i, j, arrParts, intCount

	' verify parameters are valid
	If IsNull(arrArray) Or IsEmpty(arrArray) Then
		reporter.ReportEvent micFail, strStep, "Invalid parameter: array"
		Exit Sub
	End If
	If Not IsArray(arrArray) Then
		reporter.ReportEvent micFail, strStep, "Parameter is not an Array"
		Exit Sub
	End If

	If NumberOfDimensions(arrArray) = 1 Then 
		' 1D arrays (a.k.a. single dimensional list)
		For i = 0 To Ubound(arrArray, 1)
			If CBoolean(isPurgeSeconds) Then
				If IsDate(arrArray(i)) And (InStr(arrArray(i), ":") > 0)  Then		' strip out the seconds from date/time columns
					arrParts = Split(arrArray(i), ":")
					If CBoolean(isString) Then
						arrArray(i) = arrParts(0) & ":" & arrParts(1) & " " &  Right(arrParts(2),2)
					Else
						arrArray(i) = CDate(arrParts(0) & ":" & arrParts(1) & " " &  Right(arrParts(2),2))
					End If
				End If				
			End If		
			If CBoolean(isPurgeTime) Then
				If IsDate(arrArray(i)) And (InStr(arrArray(i), ":") > 0)  Then		' strip out the time from date/time columns
					arrParts = Split(arrArray(i), " ")
					If CBoolean(isString) Then
						arrArray(i) = arrParts(0)
					Else
						arrArray(i) = CDate(arrParts(0))
					End If
				End If				
			ElseIf CBoolean(isPurgeDate) Then
				If IsDate(arrArray(i)) And (InStr(arrArray(i), ":") > 0)  Then		' strip out the date from date/time columns
					arrParts = Split(arrArray(i), " ")
					If CBoolean(isString) Then
						arrArray(i) = arrParts(1) & " " &  Right(arrParts(2),2)
					Else
						arrArray(i) = CDate(arrParts(1) & " " &  Right(arrParts(2),2))
					End If
				End If				
			End If		
		Next
	Else	
		' 2D arrays
		For i = 0 To Ubound(arrArray, 1)
			For j = 0 To Ubound(arrArray, 2)
				If CBoolean(isPurgeSeconds) Then
					If IsDate(arrArray(i, j)) And (InStr(arrArray(i, j), ":") > 0) Then		' strip out the seconds from date/time columns
						arrParts  = Split(arrArray(i, j), ":")
						If CBoolean(isString) Then
							arrArray(i, j) = arrParts(0) & ":" & arrParts(1) & " " &  Right(arrParts(2),2)
						Else
							arrArray(i, j) = CDate(arrParts(0) & ":" & arrParts(1) & " " &  Right(arrParts(2),2))
						End If
					End If
				End If			
				If CBoolean(isPurgeTime) Then
					If IsDate(arrArray(i, j)) And (InStr(arrArray(i, j), ":") > 0)  Then		' strip out the time from date/time columns
						arrParts = Split(arrArray(i, j), " ")
						If CBoolean(isString) Then
							arrArray(i, j) = arrParts(0)
						Else					
							arrArray(i, j) = CDate(arrParts(0))
						End If
					End If				
				ElseIf CBoolean(isPurgeDate) Then
					If IsDate(arrArray(i, j)) And (InStr(arrArray(i, j), ":") > 0)  Then		' strip out the date from date/time columns
						arrParts = Split(arrArray(i, j), " ")
						If CBoolean(isString) Then
							arrArray(i, j) = arrParts(1) & " " &  Right(arrParts(2),2)
						Else
							arrArray(i, j) = CDate(arrParts(1) & " " &  Right(arrParts(2),2))
						End If
					End If				
				End If		
			Next
		Next	
	End If
	
	reporter.ReportEvent micDone, strStep, "Purge date/time from array"
End Sub


''' *************************************
''' **** RegEx Functions/Subroutines ****
''' *************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Regular expression search against a specified buffer
'  strPattern = Regular expression pattern to search for
'  strBuffer = Buffer or string to be searched
'  isCaseIgnored = TRUE for case-insensitive search; otherwise FALSE for case-sensitive search
' RETURNS: TRUE if a pattern match was found; otherwise FALSE if no match was found
' NOTE: Does not log any info to result report
Function RegExpTest(ByVal strPattern, ByVal strBuffer, ByVal isCaseIgnored)
	Dim regEx
	Set regEx = New RegExp
	
	regEx.Pattern = strPattern
	regEx.IgnoreCase = isCaseIgnored
	RegExpTest = regEx.Test(strBuffer)
	
	Set regEx = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Replace text in a specified buffer using regular expression search
'  strPattern = Regular expression pattern to search for
'  strBuffer = Buffer or string to be searched and modified
'  strReplacementText = Text that will replace each pattern match
'  isCaseIgnored = TRUE for case-insensitive search; otherwise FALSE for case-sensitive search
' RETURNS: Buffer with replaced text
' NOTE: Does not log any info to result report
Function RegExpReplacement(ByVal strPattern, ByVal strBuffer, ByVal strReplacementText, ByVal isCaseIgnored)
	Dim regEx
	Set regEx = New RegExp
	
	regEx.Pattern = strPattern
	regEx.IgnoreCase = isCaseIgnored
	RegExpReplacement = regEx.Replace(strBuffer, strReplacementText)
	
	Set regEx = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get the number of regular expression matches in a specified buffer
'  strPattern = Regular expression pattern to search for
'  strBuffer = Buffer or string to be searched
'  isCaseIgnored = TRUE for case-insensitive search; otherwise FALSE for case-sensitive search
' RETURNS: Number of pattern matches found in the buffer
' NOTE: Does not log any info to result report
Function RegExpCount(ByVal strPattern, ByVal strBuffer, ByVal isCaseIgnored)
	Dim regEx, colMatches
	Set regEx = New RegExp
	
	regEx.Pattern = strPattern
	regEx.IgnoreCase = isCaseIgnored
	regEx.Global = True
	Set colMatches = regEx.Execute(strBuffer)
	RegExpCount = colMatches.Count	
	
	Set colMatches = Nothing
	Set regEx = Nothing
End Function


''' *******************************************
''' **** Folder/File Functions/Subroutines ****
''' *******************************************

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Moves a file from one location to another Or renames a file.
'  strSource = The path and filename of the original file to be moved
'  strDestination = The path (and optionally filename) where the file is to be moved. If destination is a folder is must end with '\'.
' NOTE: Does not log any info to result report
Sub MoveFile(ByVal strSource, ByVal strDestination)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject") 
	
	' move file to destination 
	If fso.FileExists(strSource) Then 
		fso.MoveFile strSource, strDestination
	End If

	Set fso = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Copies a file from one location to another
'  strSource = The path and filename of the original file to be copied
'  strDestination = The path (and optionally filename) where the file is to be moved. If destination is a folder is must end with '\'.
'  overwrite = TRUE if any existing file at the destination is to be overwritten; otherwise FALSE to never overwrite any destination file
' NOTE: Does not log any info to result report
Sub CopyFile(ByVal strSource, ByVal strDestination, ByVal overwrite)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject") 
	
	' copy file to destination 
	If fso.FileExists(strSource) Then 
		fso.CopyFile strSource, strDestination, overwrite
	End If
	
	Set fso = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Reads the contents of a specfied file
'  strFilename = The path and filename of the file to read
' RETURN: ASCII buffer with contents of the specified file
' NOTE: Does not log any info to result report
Function ReadFile(ByVal strFilename)
	Dim fso, file
	Set fso = CreateObject("Scripting.FileSystemObject") 
	Set file = fso.OpenTextFile(strFilename, 1) ' 1 = ForReading
	
	ReadFile = file.ReadAll
	file.Close
	Set file = Nothing
	Set fso = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Writes the specified contents to a file
'  strFilename = The path and filename of the file to write to
'  strContents = Contents to write
' NOTE: Does not log any info to result report
Sub WriteFile(ByVal strFilename, ByVal strContents)
	Dim fso, file
	Set fso = CreateObject("Scripting.FileSystemObject") 
	Set file = fso.CreateTextFile(strFilename, True)
	
	file.Write strContents
	file.Close
	Set file = Nothing
	Set fso = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Appends the specified contents to a file
'  strFilename = The path and filename of the file to append to
'  strContents = Contents to append
' NOTE: Does not log any info to result report
Sub AppendFile(ByVal strFilename, ByVal strContents)
	Const ForAppending = 8
	Dim fso, file
	Err.Number = 0
	Set fso = CreateObject("Scripting.FileSystemObject") 
	Set file = fso.OpenTextFile(strFilename, ForAppending, True)
	
	file.WriteLine strContents
	file.Close
	If Err.Number <> 0 Then 
		MsgBox "AppendFile failed: " & strFilename
		MsgBox ("Error Number is = " & Err.Number & "strContents attempted - " & strContents)
	End If 
	Set file = Nothing
	Set fso = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Deletes a file 
'  strFilename = The path and filename of the file to be deleted
' NOTE: Does not log any info to result report
Sub DeleteFile(ByVal strFilename)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject") 
	
	'check for existence of file. If strFilename does not exist, fso.DeleteFile would throw a fatal error, hence the inclusion of this conditional and the omission of an else clause
	If fso.FileExists(strFilename) Then
		' delete file 
		fso.DeleteFile(strFilename)
	End If
	
	
	Set fso = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Get number of files in folder
'  strFolderPath = The path of the folder containing the files
'  isIncludeSubFolders = TRUE to include all subfolders; otherwise FALSE
' RETURN: Number of files in folder; -1 if folder does not exist
' NOTE: Does not log any info to result report
Function FileCount(ByVal strFolderPath, ByVal isIncludeSubFolders)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject") 
	FileCount = 0

	If Not fso.FolderExists(strFolderPath) Then
		Set fso = Nothing
		Exit Function
	End If
	
	Dim objFolder
	Set objFolder = fso.GetFolder(strFolderPath)

	FileCount = objFolder.Files.Count
	
	If isIncludeSubFolders Then	' recursively search subfolders
		Dim objSubFolder
		
		For Each objSubFolder In objFolder.SubFolders
			FileCount = FileCount + FileCount(objSubFolder.Path, isIncludeSubFolders)
		Next
	End If
	
	Set fso = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify if folder path exists
'  strFolderPath = The path of the folder to verify
' RETURN: TRUE if folder exists; otherwise FALSE if folder does not exist
' NOTE: Does not log any info to result report
Function IsFileExists(ByVal strFileName)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject") 

	IsFileExists = fso.FileExists(strFileName)

	Set fso = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Verify if folder path exists
'  strFolderPath = The path of the folder to verify
' RETURN: TRUE if folder exists; otherwise FALSE if folder does not exist
' NOTE: Does not log any info to result report
Function IsFolderExists(ByVal strFolderPath)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject") 

	IsFolderExists = fso.FolderExists(strFolderPath)

	Set fso = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Creates a folder with the specified path
'  strFolderPath = The path of the folder to be created
' RETURN: TRUE if folder exists; otherwise FALSE if folder does not exist
' NOTE: Does not log any info to result report
Function CreateFolder(ByVal strFolderPath)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject") 

	If Not fso.FolderExists(strFolderPath) Then
		fso.CreateFolder strFolderPath		
	End If

	CreateFolder = IsFolderExists(strFolderPath)

	Set fso = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Deletes a folder
'  strFolderPath = The path of the folder to be deleted
' NOTE: Does not log any info to result report
Sub DeleteFolder(ByVal strFolderPath)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject") 
	
	If fso.FolderExists(strFolderPath) Then
		fso.DeleteFolder strFolderPath, True
	End If
	
	Set fso = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Copies contents of source folder to destination folder
'  strSource = The path of the source folder to be copied from
'  strDestination = The path of the destination folder to be copied to
'  clearDestination = TRUE if destination folder should be cleared of all files before copy process continues; otherwise FALSE to leave destination folder as-is
'  overwrite = TRUE to overwrite any existing files in the destination folder; otherwise FALSE to leave existing destination files as-is
' RETURN: TRUE if folder exists; otherwise FALSE if folder does not exist
' NOTE: Does not log any info to result report
Function CopyFolder(ByVal strSource, ByVal strDestination, ByVal clearDestination, ByVal overwrite)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject") 

	If clearDestination And fso.FolderExists(strDestination) Then
		fso.DeleteFolder strDestination, True
	End If

'	If Not fso.FolderExists(strDestination) Then
'		fso.CreateFolder strDestination
'	End If

    fso.CopyFolder strSource, strDestination, overwrite
	CopyFolder = IsFolderExists(strDestination)

	Set fso = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Moves the source folder to destination folder
'  strSource = The path of the source folder to move from
'  strDestination = The path of the destination folder to move to
' RETURN: TRUE if folder exists; otherwise FALSE if folder does not exist
' NOTE: Does not log any info to result report
Function MoveFolder(ByVal strSource, ByVal strDestination)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject") 

	If fso.FolderExists(strDestination) Then
		fso.DeleteFolder strDestination, True
	End If

'	If Not fso.FolderExists(strDestination) Then
'		fso.CreateFolder strDestination
'	End If

    fso.MoveFolder strSource, strDestination
	MoveFolder = IsFolderExists(strDestination)

	Set fso = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Builds the full path for a specified path and file/folder name (inserts delimiter, if necessary)
'  strPath = The path to the folder
'  strName = Name of file or folder to append
' RETURN: Full path with path delimiters
' NOTE: Does not log any info to result report
Function BuildPath(ByVal strPath, ByVal strName)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject") 

	BuildPath = fso.BuildPath( strPath, strName)

	Set fso = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Gets the local path for the current user's MyDocuments folder
' RETURN: Path to current user's MyDocuments folder
' NOTE: Does not log any info to result report
Function GetMyDocumentsPath()
	Dim fso, file
	Set fso = CreateObject("Scripting.FileSystemObject") 

	GetMyDocumentsPath = CreateObject("Wscript.Shell").SpecialFolders("Mydocuments")

	Set fso = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Gets the local path for the current project
' RETURN: Path to current project
' NOTE: Does not log any info to result report
Function ProjectWorkingPath()
	ProjectWorkingPath =  Left(Environment("TestDir"), Len(Environment("TestDir"))-Len(Environment("TestName")))
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DESC: Append entry to a Results log file
'  strFile = The path and filename of the results file (Null = disable logging)
'  boolIsPassed = TRUE to log Passed; otherwise FALSE to log Failed
'  dtStartTime = Start time of test
'  dtStartTime = End time of test
'  strName = Name of test (Null = default:N/A)
'  strReference = Test step Reference text (Null = default:N/A)
'  strDescription = Test step Description text (Null = default:N/A)
'  strExpectedResult = Test step Expected Result text (Null = default:N/A)
' NOTE: Does not log any info to UFT result report
Sub LogResult(ByVal strFile, ByVal boolIsPassed, ByVal dtStartTime, ByVal dtEndTime, ByVal strName, ByVal strReference, ByVal strDescription, ByVal strExpectedResult)
	If StringIsNullOrEmpty(strFile) Then
		Exit Sub
	End If
	If StringIsNullOrEmpty(dtStartTime) Then
		dtStartTime = Now()
	End If
	If StringIsNullOrEmpty(dtEndTime) Then
		dtEndTime = Now()
	End If
	If StringIsNullOrEmpty(strName) Then
		strName = "N/A"
	End If
	If StringIsNullOrEmpty(strReference) Then
		strReference = "N/A"
	End If
	If StringIsNullOrEmpty(strDescription) Then
		strDescription = "N/A"
	End If
	If StringIsNullOrEmpty(strExpectedResult) Then
		strExpected = "N/A"
	End If

	Dim strType : strType = LCase(Right(strFile,3))	' file extension defines type of logging
	
	If strType = "csv" Then
		AppendFile strFile, dtStartTime & "," & dtEndTime & "," & boolIsPassed & "," & strName & "," & strReference  & "," & strDescription  & "," & strExpectedResult
	ElseIf strType = "txt" Then
		AppendFile strFile, dtStartTime & "#" & dtEndTime & "#" & boolIsPassed & "#" & strName & "#" & strReference  & "#" & strDescription  & "#" & strExpectedResult
	ElseIf strType = "xls" Then
		On Error Resume Next
		Dim strLogSheet : strLogSheet = "RESULTS"
		With DataTable		
			.GetSheet(strLogSheet).SetCurrentRow(datatable.GetSheet(strLogSheet).GetRowCount + 1)
			.Value("Test", strLogSheet) = strName
			.Value("Reference", strLogSheet) = strReference
			.Value("Description", strLogSheet) = strDescription
			.Value("ExpectedResult", strLogSheet) = strExpectedResult
			.Value("StartTime", strLogSheet) = dtStartTime
			.Value("EndTime", strLogSheet) = dtEndTime
			.Value("Passed", strLogSheet) = boolIsPassed
		End With
	ElseIf LCase(Right(strFile,3)) = "xml" Then
		' NOT IMPLEMENTED YET: FOR FUTURE USE
	Else
		MsgBox "LogResult failed - Unsupported file format", "Error"
	End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' END SCRIPT
