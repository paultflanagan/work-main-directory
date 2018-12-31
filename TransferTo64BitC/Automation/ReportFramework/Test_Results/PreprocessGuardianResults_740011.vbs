Option Explicit

Dim FSO, FileInput, FileOutput, WorkSheetArray, FrameWorkFolderPath

Redim WorkSheetArray(7, -1)	' column/row zero-based

' read summary data
ReadDataIntoArray False, "GuardianSummary_740011.txt", WorkSheetArray
' read detail data
ReadDataIntoArray True, "Results.txt", WorkSheetArray

Set FSO = CreateObject("Scripting.FileSystemObject")
Set FileOutput = FSO.CreateTextFile("Project_Report.txt", True)

' write combined data
SaveArrayDataToCSV FileOutput, WorkSheetArray

FileOutput.Close
Set FileOutput = Nothing

'FrameWorkFolderPath ="C:\jenkins\workspace\Data_Integrity_Tool_Automation\Project_Framework"
'FSO.CopyFile "Project_Report.txt", FrameWorkFolderPath & "\Test_Results\Project_Report.txt"

Set FSO = Nothing

'
'----------------------------------------------------------------------------------------------------------------------------
Public Sub ReadDataIntoArray(ByVal IsDetail, ByVal InputFile, ByRef ArrayList)
	Dim FSO, FileOpenObj, Count, RowCount, i, LineValue, TempArrayList

	RowCount = UBound(ArrayList, 2)
	
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Set FileOpenObj = FSO.OpenTextFile(InputFile, 1)
	
	Count = 0
	Do Until FileOpenObj.AtEndOfStream
		FileOpenObj.Readline 'Prints each line until the whole text file is read
		Count = Count + 1
	Loop
	
	FileOpenObj.Close
	Set FileOpenObj = Nothing

	Redim Preserve ArrayList(UBound(ArrayList, 1), RowCount + Count)
	Set FileOpenObj = FSO.OpenTextFile(InputFile, 1)
	
	For i = (RowCount + 1) to UBound(ArrayList, 2)
		LineValue = FileOpenObj.Readline
		TempArrayList = Split(LineValue, "#")
		If IsDetail Then
			ArrayList(0,i) = i - RowCount
			ArrayList(1,i) = Trim(TempArrayList(3))
			ArrayList(2,i) = Trim(TempArrayList(4))
			ArrayList(3,i) = Trim(TempArrayList(5))
			ArrayList(4,i) = Trim(TempArrayList(6))
			ArrayList(6,i) = convertTime(DateDiff("s",Trim(TempArrayList(0)),Trim(TempArrayList(1))))
			
			If CBool(Trim(TempArrayList(2))) Then
				ArrayList(7,i) = "PASS"
			Else
				ArrayList(7,i) = "FAIL"
			End If
			
			If i = (RowCount + 1) Then					' first start time
				ArrayList(5,0) = Trim(TempArrayList(0))	' session start time
			End If
			If i = UBound(ArrayList, 2) Then			' last end time
				ArrayList(5,1) = Trim(TempArrayList(1))	' session end time
				ArrayList(5,2) = convertTime(DateDiff("s", ArrayList(5,0), ArrayList(5,1)))
			End If
		Else ' header summary
			ArrayList(0,i) = Trim(TempArrayList(0))
			ArrayList(1,i) = Trim(TempArrayList(1))
			ArrayList(2,i) = Trim(TempArrayList(2))
			ArrayList(3,i) = Trim(TempArrayList(3))
			ArrayList(4,i) = Trim(TempArrayList(4))
			ArrayList(5,i) = Trim(TempArrayList(5))
			ArrayList(6,i) = Trim(TempArrayList(6))
			ArrayList(7,i) = Trim(TempArrayList(7))
		End If
	Next
	
	FileOpenObj.Close
	Set FileOpenObj = Nothing
End Sub

'----------------------------------------------------------------------------------------------------
Public Sub SaveArrayDataToCSV(FileOpenObj, TwoDimArrayData)
	Dim i
	
	For i = 0 to UBOUND(TwoDimArrayData, 2)
		If i < UBOUND(TwoDimArrayData, 2) Then
			FileOpenObj.Writeline TwoDimArrayData(0, i) & "#" & TwoDimArrayData(1, i) & "#" & TwoDimArrayData(2, i) & "#" & TwoDimArrayData(3, i) & "#" _
			& TwoDimArrayData(4, i) & "#" & TwoDimArrayData(5, i) & "#" & TwoDimArrayData(6, i) & "#" & TwoDimArrayData(7, i)
		Else
			FileOpenObj.Write TwoDimArrayData(0, i) & "#" & TwoDimArrayData(1, i) & "#" & TwoDimArrayData(2, i) & "#" & TwoDimArrayData(3, i) & "#" _
			& TwoDimArrayData(4, i) & "#" & TwoDimArrayData(5, i) & "#" & TwoDimArrayData(6, i) & "#" & TwoDimArrayData(7, i)
		End If
	Next
End Sub

'----------------------------------------------------------------------------------------------------
Public Function convertTime(intSeconds)
	Dim strSec, strMin, strHour
	
	strSec = intSeconds Mod 60
	If Len(strSec) = 1 Then
		 strSec = "0" & strSec
	End If
	
	strMin = (intSeconds Mod 3600) \ 60
	If Len(strMin) = 1 Then
		 strMin = "0" & strMin
	End If
	
	strHour =  intSeconds \ 3600
	If Len(strHour) = 1 Then
		 strHour = "0" & strHour
	End If
	
	convertTime = strHour & ":" & strMin & ":" & strSec
End Function
'----------------------------------------------------------------------------------------------------

