'FrameWorkFolderPath = "C:Automation\Duplicate Check\ReportFramework"

Set WscriptObj = CreateObject("WScript.shell")
'DriverScriptFrameWorkFolderPath = WscriptObj.currentdirectory
'This was not a robust enough way of accessing C:\Automation\ReportFramework\Test_Driver. Depending on where the Window's focus was, this command would occasionally result in pointing to system32
'Hard coding solution now, fix later to be dynamic
DriverScriptFrameWorkFolderPath = "C:\Automation\ReportFramework\Test_Driver"
FrameWorkFolderPath = LEFT(DriverScriptFrameWorkFolderPath, (InStrRev(DriverScriptFrameWorkFolderPath, "\") - 1))
Set WscriptObj = Nothing

Set FSO = CreateObject("Scripting.FileSystemObject")
Set FilesObj = FSO.GetFolder(FrameWorkFolderPath &"\Test_Results").Files

If FSO.FileExists(FrameWorkFolderPath &"\Test_Results\Project_Report_duplicatecheck.txt") Then
	Call FSO.DeleteFile(FrameWorkFolderPath &"\Test_Results\Project_Report_duplicatecheck.txt")
End If

Count = 0
Redim NameArrayList(FilesObj.Count - 1)

For Each Files in FilesObj
	If FindFileExtension(Files.Name) = "txt" then
		If INSTR(lcase(Files.Name), "project_report") > 0 Then
			NameArrayList(Count) = Files.Name
			Count = Count + 1
		End If
	End If
Next

SortedNameArrayList = CombinedArraySort(NameArrayList, "DESC")
Call FSO.CopyFile(FrameWorkFolderPath &"\Test_Results\" & SortedNameArrayList(0), FrameWorkFolderPath &"\Test_Results\Project_Report_duplicatecheck.txt")
'msgbox("Creation attempt complete.")
'Call FSO.CopyFile(FrameWorkFolderPath &"\Test_Results\Project_Report_duplicatecheck.txt", "Y:\jenkins\workspace\Automation\UniSeries\8.4.0\SVT\Duplicate Check Report\Project_Report_duplicatecheck.txt")
'Call FSO.CopyFile(FrameWorkFolderPath &"\Test_Results\Project_Report_duplicatecheck.txt", "Y:\Project_Report_duplicatecheck.txt")
Set FSO = Nothing

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