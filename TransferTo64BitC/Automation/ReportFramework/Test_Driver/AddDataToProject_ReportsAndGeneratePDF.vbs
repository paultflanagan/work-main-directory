'FrameWorkFolderPath ="C:\jenkins\workspace\Data_Integrity_Tool_Automation\Project_Framework"
TestEnvironmentSheetName = "TestEnvironmentSheet.txt"

Set WscriptObj = CreateObject("WScript.shell")
DriverScriptFolderPath = WscriptObj.currentdirectory
FrameWorkFolderPath = LEFT(DriverScriptFolderPath, InStrRev(DriverScriptFolderPath, "\") - 1)
Set WscriptObj = Nothing 

Set FSO = CreateObject("Scripting.FileSystemObject")
Set FilesObj = FSO.GetFolder(FrameWorkFolderPath &"\Test_Results").Files

Count = 0
Redim NameArrayList(FilesObj.Count - 1)

For Each Files in FilesObj
   	If FindFileExtension(Files.Name) = "txt" then
		If INSTR(lcase(Files.Name), "project_report") > 0 Then
		'If INSTR(lcase(Files.Name), "project_report_unisphere") > 0 Then
			ProjectFileName = Files.Name
			'NameArrayList(Count) = Files.Name
			'Count = Count + 1
		End If
	End If
Next

'If isEmpty(NameArrayList(0)) = False Then
If isEmpty(ProjectFileName) = False Then
	'SortedNameArrayList = CombinedArraySort(NameArrayList, "DESC")
	'ProjectReportWorkSheetArray = ConvertCSVDataInto2DArray(FrameWorkFolderPath &"\Test_Results\" &SortedNameArrayList(0), "TestCaseStatistics")
	ProjectReportWorkSheetArray = ConvertCSVDataInto2DArray(FrameWorkFolderPath &"\Test_Results\" &ProjectFileName, "TestCaseStatistics")
	ProjectReportWorkSheetArray(0,7) = "=COUNTIF(H7:H190, ""PASS"")"
	ProjectReportWorkSheetArray(1,7) = "=COUNTIF(H7:H190, ""FAIL"")"
	ProjectReportWorkSheetArray(2,7) = "=COUNTIF(H7:H190, ""SKIPPED"")"
	ProjectReportWorkSheetArray(3,7) = "=SUM(H1:H3)"
	ProjectReportWorkSheetArray(3,5) = "=IF(H2 > 0, ""FAIL"", IF(H1 > 0, ""PASS"", IF(H3 > 0, ""SKIPPED"", """")))"
	'ProjectFileNameFirstPart = Replace(SortedNameArrayList(0), ".txt", "")
	ProjectFileNameFirstPart = Replace(ProjectFileName, ".txt", "")
	Call FSO.CopyFile(FrameWorkFolderPath &"\Test_Driver\Project_ReportHSV.xlsx", FrameWorkFolderPath &"\Test_Results\" &ProjectFileNameFirstPart &".xlsx")
	Call SaveArrayDataToExcel(FrameWorkFolderPath &"\Test_Results\" &ProjectFileNameFirstPart &".xlsx", "TestCaseStatistics", ProjectReportWorkSheetArray)
	'Call FSO.DeleteFile(FrameWorkFolderPath &"\Test_Results\" &SortedNameArrayList(0))
	Call FSO.DeleteFile(FrameWorkFolderPath &"\Test_Results\" &ProjectFileName)
	Set XLObj = CreateObject("Excel.Application")
	Set WBObj = XLObj.Workbooks.Open(FrameWorkFolderPath &"\Test_Results\" &ProjectFileNameFirstPart &".xlsx")
	Call WBObj.ExportAsFixedFormat (0, FrameWorkFolderPath &"\Test_Results\" &ProjectFileNameFirstPart, 0, True, False, , , False)
	Wscript.Sleep(3000)
	WBObj.Close
	Set WBObj = Nothing
	'Call FSO.DeleteFile(FrameWorkFolderPath &"\Test_Results\" &ProjectFileNameFirstPart &".xlsx")
	Set FilesObj = Nothing
	Set FSO = Nothing
End If

'
'----------------------------------------------------------------------------------------------------------------------------
Public Function ConvertCSVDataInto2DArray(FolderPathAndFilename, SheetName)
	Dim FSO, FileOpenObj, Count, RowCount, ColumnCount, ArrayList, I, J, LineValue, TempArrayList
	
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set FileOpenObj = FSO.OpenTextFile(FolderPathAndFilename, 1)
	
	DO UNTIL FileOpenObj.AtEndOfStream
		FileOpenObj.Readline 'Prints each line until the whole text file is read
		Count = Count + 1
	LOOP
	
	FileOpenObj.Close
	Set FileOpenObj = Nothing

	RowCount = Count
	
	If SheetName = "TestPlan" Then 
		ColumnCount = 11
	Else
		If SheetName = "TestLab" Then 
			ColumnCount = 4
		ElseIf SheetName = "TestEnvironmentValues" Then
			ColumnCount = 3
		ElseIf SheetName = "TestCaseStatistics" Then
			ColumnCount = 8
		End If
	End If
	
	Redim ArrayList(RowCount - 1, ColumnCount - 1)
	Set FileOpenObj = FSO.OpenTextFile(FolderPathAndFilename, 1)
	
	DO UNTIL FileOpenObj.AtEndOfStream
		For I = 0 to RowCount - 1
			LineValue = FileOpenObj.Readline
			For J = 0 to ColumnCount - 1
				TempArrayList = Split(LineValue, "#")
				ArrayList(I, J) = Trim(TempArrayList(J))
				Print ArrayList(I, J)
			Next
		Next
	LOOP
	FileOpenObj.Close
	ConvertCSVDataInto2DArray = ArrayList
End Function
'----------------------------------------------------------------------------------------------------
'
'----------------------------------------------------------------------------------------------------
Public Sub SaveArrayDataFromCSVToProjectReportExcel(FolderPathAndFilename, SheetName, TwoDimArrayData)
	Dim ArrayXLObj, ArrayWBObj, XLWSObj, RowCount, ColumnCount, I, J
	
	Set ArrayXLObj = CreateObject("Excel.Application")
	Set ArrayWBObj = ArrayXLObj.Workbooks.Open(FolderPathAndFilename)
	Set XLWSObj = ArrayWBObj.Sheets(SheetName)
	RowCount = UBOUND(TwoDimArrayData, 1)
	ColumnCount = UBOUND(TwoDimArrayData, 2)
	
	For I = 0 to RowCount
		For J = 0 to ColumnCount
			XLWSObj.Cells(I + 1, J + 1).Formula = TwoDimArrayData(I, J) 
		Next
	Next
	ArrayWBObj.Save
	Set XLWSObj = Nothing
	Set ArrayWBObj = Nothing
	ArrayXLObj.Quit
	Set ArrayXLObj = Nothing
End Sub
'--------------------------------------------------------------------------------
Public Function CombinedArraySort(ArrayList, SortOrder)
	Dim NumericCount, StringCount, I, Num1, Str1
	
	NumericCount = 0 : StringCount = 0

	For I = 0 to UBOUND(ArrayList)
		If isNumeric(ArrayList(I)) then 
			NumericCount = NumericCount + 1
		Else 
			StringCount = StringCount + 1
		End If
	Next
	
	If NumericCount = 0 Then
		CombinedArraySort = ArraySort(ArrayList, SortOrder)
	Else
		If StringCount = 0 Then 
			Redim NumericArrayList(NumericCount - 1)
			Num1 = 0
				For I = 0 to UBOUND(ArrayList)
					NumericArrayList(Num1) = Cint(ArrayList(I))
					Num1 = Num1 + 1
				Next
			CombinedArraySort = ArraySort(NumericArrayList, SortOrder)
		Else
			Redim NumericArrayList(NumericCount - 1)
			Redim StringArrayList(StringCount - 1)
			Num1 = 0 : Str1 = 0		
			For I = 0 to UBOUND(ArrayList)
				If isNumeric(ArrayList(I)) then 
					NumericArrayList(Num1) = Cint(ArrayList(I))
					Num1 = Num1 + 1
				Else 
					StringArrayList(Str1) = trim(ArrayList(I))
					Str1 = Str1 + 1
				End If
			Next
			CombinedArraySort = AddNewValuesToAnArray(ArraySort(StringArrayList, SortOrder), ArraySort(NumericArrayList, SortOrder))
		End If
	End If
End Function

Public Function ArraySort(ArrayList, SortOrder)
	Dim I, J, TempVariable
	IF TrimAndUcase(SortOrder) = "ASC" OR TrimAndUcase(SortOrder) = "ASCENDING" Then 
		For I = UBOUND(ArrayList) - 1 To 0 Step-1
			For J = 0 to I
				If ArrayList(J) > ArrayList(J + 1) Then
					TempVariable = ArrayList(J)
					ArrayList(J) = ArrayList(J + 1)
					ArrayList(J + 1) = TempVariable
				End if
			Next
		Next
		ArraySort = ArrayList
	ElseIf TrimAndUcase(SortOrder) = "DESC" OR TrimAndUcase(SortOrder) = "DESCENDING" Then 
		For I = UBOUND(ArrayList) - 1 To 0 Step-1
			For J = 0 to I
				If ArrayList(J) < ArrayList(J + 1) Then
					TempVariable = ArrayList(J)
					ArrayList(J) = ArrayList(J + 1)
					ArrayList(J + 1) = TempVariable
				End if
			Next
		Next
		ArraySort = ArrayList
	Else
		ArraySort = "Invalid SortOrder, please specify either of the following ASC, DESC, ASCENDING or DESCENDING"
	End If 
End Function

Public Function TrimAndUcase(StringValue)
	TrimAndUcase = UCASE(TRIM(StringValue))
End Function

Public Function AddNewValuesToAnArray(OriginalArray, NewArrayWithNewValues)
	Dim NewBound, Count, I, J
	NewBound = UBOUND(OriginalArray) + UBOUND(NewArrayWithNewValues) + 1
	Redim NewArrayList(NewBound)
	Count = 0 
	FOR I = 0 to UBOUND(OriginalArray)
		NewArrayList(Count) = OriginalArray(I)
		Count = Count + 1
	Next
	
	For J = 0 to UBOUND(NewArrayWithNewValues)
		NewArrayList(Count) = NewArrayWithNewValues(J)
		Count = Count + 1
	Next
	AddNewValuesToAnArray = NewArrayList
End Function

Public Function FindFileExtension(FileNameOrFolderPathWithFileName)
	IF LEFT(RIGHT(FileNameOrFolderPathWithFileName, 4), 1) = "." Then
		FindFileExtension = RIGHT(FileNameOrFolderPathWithFileName, 3)
	ElseIF LEFT(RIGHT(FileNameOrFolderPathWithFileName, 5), 1) = "." Then
		FindFileExtension = RIGHT(FileNameOrFolderPathWithFileName, 4)
	Else 
		FindFileExtension = "Invalid File name or File path"
	End if 
End Function


'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ConvertCSVDataInto2DArray(FolderPathAndFilename, SheetName)
	Dim FSO, FileOpenObj, Count, RowCount, ColumnCount, ArrayList, I, J, LineValue, TempArrayList, TempVariable
	
	Set FSO = CreateObject("Scripting.FileSystemObject")
	'Print "FolderPathAndFilename - " &FolderPathAndFilename
	Set FileOpenObj = FSO.OpenTextFile(FolderPathAndFilename, 1)
	
	DO UNTIL FileOpenObj.AtEndOfStream
		TempVariable = FileOpenObj.Readline 'Prints each line until the whole text file is read
		If Len(TempVariable) > 0 Then
			Count = Count + 1
		End If
	LOOP
	
	FileOpenObj.Close
	Set FileOpenObj = Nothing

	RowCount = Count
	
	If SheetName = "TestPlan" Then 
		ColumnCount = 11
	Else
		If SheetName = "TestLab" Then 
			ColumnCount = 4
		ElseIf SheetName = "TestEnvironmentValues" Then
			ColumnCount = 3
		ElseIf SheetName = "TestCaseStatistics" Then
			ColumnCount = 8
		End If
	End If
	
	Redim ArrayList(RowCount - 1, ColumnCount - 1)
	Set FileOpenObj = FSO.OpenTextFile(FolderPathAndFilename, 1)
	
	DO UNTIL FileOpenObj.AtEndOfStream
		For I = 0 to RowCount - 1
			LineValue = FileOpenObj.Readline
			'If I = 4 Then
				For J = 0 to ColumnCount - 1
					TempArrayList = Split(LineValue, "#")
					ArrayList(I, J) = Trim(TempArrayList(J))
				Next
		'	Else
				'For J = 0 to ColumnCount - 1
				'	TempArrayList = Split(LineValue, "#")
				'	ArrayList(I, J) = Trim(TempArrayList(J))
			'	Next
			'End If
		Next
	LOOP
	FileOpenObj.Close
	Set FileOpenObj = Nothing
	ConvertCSVDataInto2DArray = ArrayList
End Function
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub SaveArrayDataToExcel(FolderPathAndFilename, SheetName, TwoDimArrayData)
	Dim ArrayXLObj, ArrayWBObj, XLWSObj, RowCount, ColumnCount, I, J, K, TempArray, TempCellValue
	
	Set ArrayXLObj = CreateObject("Excel.Application")
	Set ArrayWBObj = ArrayXLObj.Workbooks.Open(FolderPathAndFilename)
	Set XLWSObj = ArrayWBObj.Sheets(SheetName)
	RowCount = UBOUND(TwoDimArrayData, 1)
	ColumnCount = UBOUND(TwoDimArrayData, 2)
	
	For I = 0 to RowCount
		For J = 0 to ColumnCount
			If J = 3 OR J = 4 Then
					TempArray = Split(TwoDimArrayData(I, J), "@")
					For K = 0 to UBOUND(TempArray)
						If K = UBOUND(TempArray) Then
							TempCellValue = TempCellValue &TempArray(K)
						Else						
							TempCellValue = TempCellValue &TempArray(K) &vbLf
						End If
					Next
				
				XLWSObj.Cells(I + 1, J + 1).Formula = TempCellValue
				TempCellValue = ""
			Else
				XLWSObj.Cells(I + 1, J + 1).Formula = TwoDimArrayData(I, J) 
			End If
		Next
	Next
	ArrayWBObj.Save
	Set XLWSObj = Nothing
	Set ArrayWBObj = Nothing
	ArrayXLObj.Quit
	Set ArrayXLObj = Nothing
End Sub
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ExtractEachElementInListToArray(StringValue)
	Dim RegExpObj, MatchCollectionObj, ArrayList, I, EachMatchItem
	Set RegExpObj = New RegExp
	With RegExpObj
		.Pattern = ".+\S+"
		.IgnoreCase = False
		.Global = True
	End With

	Set MatchCollectionObj = RegExpObj.Execute(StringValue)
	Redim ArrayList(MatchCollectionObj.Count-1)

	I = 0
	For Each EachMatchItem in MatchCollectionObj
		ArrayList(I) = EachMatchItem.Value
		I = I + 1
	Next
	ExtractEachElementInListToArray = ArrayList
	Set RegExpObj = Nothing
End Function
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Function ExtractEachElementInListToArrayMatchingPattern(StringValue, PatternMatch)
	Dim RegExpObj, MatchCollectionObj, ArrayList, I, EachMatchItem
	Set RegExpObj = New RegExp
	With RegExpObj
		.Pattern = PatternMatch
		.IgnoreCase = False
		.Global = True
	End With

	Set MatchCollectionObj = RegExpObj.Execute(StringValue)
	Redim ArrayList(MatchCollectionObj.Count-1)

	I = 0
	For Each EachMatchItem in MatchCollectionObj
		ArrayList(I) = EachMatchItem.Value
		I = I + 1
	Next
	ExtractEachElementInListToArrayMatchingPattern = ArrayList
	Set RegExpObj = Nothing
End Function
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
Public Function CorrectMultipleLineData(StringValue)
	Set FSO = CreateObject("Scripting.FileSystemObject")
	If FSO.FileExists("c:\TEMP\Test123.txt") Then 
		Set OpenTextFile1 = FSO.OpenTextFile("c:\TEMP\Test123.txt", 2)
		OpenTextFile1.Write trim(StringValue)
		'OpenTextFile.Close
	Else
		FSO.CreateTextFile("c:\TEMP\Test123.txt")
		Set OpenTextFile1 = FSO.OpenTextFile("c:\TEMP\Test123.txt", 2)
		OpenTextFile1.Write trim(StringValue)
		'OpenTextFile.Close
	End If
		Set OpenTextFile2 = FSO.OpenTextFile("c:\TEMP\Test123.txt", 1)
		CorrectMultipleLineData = OpenTextFile2.ReadAll
End Function
'----------------------------------------------------------------------------------------------------